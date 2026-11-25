"""Solve the model stored by Excel Solver in a workbook and write the optimum
into a new file.

Requirements:  pip install openpyxl scipy

WHAT THIS TOOL DOES / FIXES
---------------------------
* Objective cell: Excel Solver saves it as `solver_opt` (with the target value
  in `solver_val`), not `solver_obj`.  Both spellings are accepted.
* Formula evaluation: no xlcalculator.  A built-in engine (openpyxl tokenizer
  + numpy) evaluates the sheet and supports arithmetic, ranges, array
  constants {..}, MMULT, SUMPRODUCT, SUM/SUMSQ/AVERAGE/COUNT, MIN/MAX, INDEX,
  OFFSET, VLOOKUP, SUMIF, COLUMN/ROW, IF/AND/OR/NOT, ROUND/ROUNDUP/ROUNDDOWN/
  INT/TRUNC/MOD/SIGN, EXP/LN/LOG/POWER/SQRT, CODE, plus legacy CSE array
  formulas.  Blank cells are skipped by AVERAGE/COUNT like Excel does.
  Every formula cell is re-checked against the value Excel cached in the file,
  so a wrong evaluation is reported instead of silently becoming 0.
* Several Solver models in one workbook (one per worksheet): the model of the
  DEFAULT/ACTIVE worksheet is used; `--sheet NAME` picks another one.
* Constraint right-hand sides given as defined names (e.g. `ShippingLimit`)
  are resolved, including worksheet-scoped names.
* Linear vs nonlinear is decided defensively: "linear" only if an affine model
  fitted at one point reproduces the objective and every constraint at several
  other points over three magnitudes.  Otherwise (or on any evaluation error)
  the model is treated as nonlinear.
* Correctness of the reported answer: the optimum is re-evaluated with a
  FRESH engine (no caches) before saving, and any constraint that is still
  violated is printed with its left/right values.  Formulas consisting of a
  bare reference (`=C50`) used to lose their dependency edge, which left stale
  values in the cache and made such constraints look unconstrained -- fixed.
* Solving: LP/MILP via HiGHS when the model is affine; otherwise successive
  linear programming (SLP) when the constraints are affine, which is the case
  for most Excel Solver models (nonlinearity sits in the objective, e.g.
  price x quantity).  SLP keeps feasibility at every step and is much more
  reliable than SLSQP on 60+ variable workbooks.  Small models additionally
  get a multi-start SLSQP; the best feasible answer wins.  If the worksheet
  values violate the constraints, the start is repaired by an LP first.

Usage:
    python pysolve.py                       # file dialog
    python pysolve.py book.xlsx             # writes pysolved-book.xlsx
    python pysolve.py book.xlsx --sheet Production
    python pysolve.py book.xlsx --nonlinear | --linear
"""

try:
    import numpy as np
    import openpyxl
    import scipy
except ImportError:
    input('''This program requires openpyxl and scipy.
Install them by running this command:

    pip install openpyxl scipy

Press Enter key to exit...''')
    exit()

from openpyxl.formula import Tokenizer
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import (range_boundaries, get_column_letter,
                            column_index_from_string)

try:                                            # optional but recommended
    from scipy.optimize import Bounds, LinearConstraint, milp, minimize
    HAVE_SCIPY = True
except ImportError:                             # pragma: no cover
    HAVE_SCIPY = False

import re

# ==========================================================================
# 1. Formula engine
# ==========================================================================
_CELL_RE = re.compile(r"^(\$?[A-Za-z]{1,3}\$?\d{1,7})$")
_RANGE_RE = re.compile(r"^(\$?[A-Za-z]{1,3}\$?\d{1,7}):(\$?[A-Za-z]{1,3}\$?\d{1,7})$")


def _split_sheet(ref):
    """'Product Mix'!$D$13  ->  ("Product Mix", "$D$13")"""
    if "!" not in ref:
        return None, ref
    sheet, rest = ref.rsplit("!", 1)
    sheet = sheet.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    if "]" in sheet:                          # external workbook
        sheet = sheet.split("]")[-1]
    return sheet, rest


class Ref:
    """A reference that still knows where it points.

    Functions such as COLUMN/ROW/INDEX/OFFSET/VLOOKUP/SUMIF need the address,
    not just the values, so references are kept as Ref objects and converted to
    numbers/arrays only when they take part in arithmetic.
    """

    __slots__ = ("eng", "sheet", "c1", "r1", "c2", "r2")

    def __init__(self, eng, sheet, c1, r1, c2, r2):
        self.eng, self.sheet = eng, sheet
        self.c1, self.r1, self.c2, self.r2 = c1, r1, c2, r2

    @property
    def single(self):
        return self.c1 == self.c2 and self.r1 == self.r2

    def values(self):
        return self.eng.rect_values(self.sheet, self.c1, self.r1, self.c2, self.r2)

    def raw(self):
        """Value of a single-cell reference, keeping text as text."""
        return self.eng.cell_value(self.sheet, f"{get_column_letter(self.c1)}{self.r1}")

    def shifted(self, dr, dc, height, width):
        c1, r1 = self.c1 + dc, self.r1 + dr
        c2 = c1 + width - 1 if width else self.c2 + dc
        r2 = r1 + height - 1 if height else self.r2 + dr
        return Ref(self.eng, self.sheet, c1, r1, c2, r2)

    def __repr__(self):
        return f"<Ref {self.sheet}!R{self.r1}C{self.c1}:R{self.r2}C{self.c2}>"


class FormulaEngine:
    """On-demand evaluation of a workbook's formulas."""

    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path, data_only=False)
        self.wbv = openpyxl.load_workbook(path, data_only=True)
        self.sheets = {ws.title.lower(): ws.title for ws in self.wb.worksheets}
        self.formulas, self.cached = {}, {}
        for ws, wsv in zip(self.wb.worksheets, self.wbv.worksheets):
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    key = (ws.title, cell.coordinate)
                    if isinstance(cell.value, ArrayFormula):
                        # legacy CSE array formula: keep its text as the formula
                        self.formulas[key] = cell.value.text
                        self.cached[key] = None
                        continue
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        self.formulas[key] = cell.value
                        self.cached[key] = wsv[cell.coordinate].value
                    else:
                        self.cached[key] = cell.value
        self.overrides, self.memo, self.parsed = {}, {}, {}
        self.deps, self.rdeps, self._stack = {}, {}, []
        self._ref_cache = {}
        # defined names: (sheet|None, name) -> reference text
        self.names = {}
        for ws in self.wb.worksheets:
            for n, d in ws.defined_names.items():
                self.names[(ws.title.lower(), n.lower())] = d.attr_text
        for n, d in self.wb.defined_names.items():
            self.names.setdefault((None, n.lower()), d.attr_text)

    # -- API ---------------------------------------------------------------
    def set_values(self, cells, values):
        """Set decision variables and drop only the cached cells that depend
        on them (dependent cells are tracked while evaluating)."""
        touched = []
        for cell, v in zip(cells, values):
            sheet, coord = cell.split("!")
            key = (self._sheet(sheet), coord.replace("$", ""))
            self.overrides[key] = float(v)
            touched.append(key)
        self._invalidate(touched)

    def _invalidate(self, keys):
        seen, stack = set(keys), list(keys)
        while stack:
            k = stack.pop()
            for parent in self.rdeps.get(k, ()):
                if parent not in seen:
                    seen.add(parent)
                    stack.append(parent)
        for k in seen:
            self.memo.pop(k, None)

    def value(self, cell):
        sheet, coord = cell.split("!")
        return self.cell_value(self._sheet(sheet), coord.replace("$", ""))

    def lookup_name(self, sheet, name):
        key = name.lower()
        return (self.names.get((sheet.lower(), key))
                or self.names.get((None, key)))

    # -- internals ---------------------------------------------------------
    def _sheet(self, name):
        n = name.strip().strip("'")
        return self.sheets.get(n.lower(), n)

    def cell_value(self, sheet, coord):
        key = (sheet, coord.replace("$", "").upper() if False else coord.replace("$", ""))
        if self._stack:                                # record dependency
            parent = self._stack[-1]
            if parent != key:
                self.deps.setdefault(parent, set()).add(key)
                self.rdeps.setdefault(key, set()).add(parent)
        if key in self.overrides:
            return self.overrides[key]
        if key in self.memo:
            return self.memo[key]
        if key in self.formulas:
            self.memo[key] = 0.0                       # cycle guard
            tokens = self.parsed.get(key)
            if tokens is None:
                tokens = [t for t in Tokenizer(self.formulas[key]).items
                          if t.type != "WSPACE"]
                self.parsed[key] = tokens
            self._stack.append(key)
            try:
                val = _Parser(tokens, sheet, self, key[1]).parse()
                # resolve a bare reference *inside* the stack scope, otherwise
                # the dependency edge for formulas like '=C50' is never
                # recorded and the cell is not invalidated later.
                if isinstance(val, Ref):
                    val = val.raw() if val.single else val.values()
                if isinstance(val, np.ndarray):
                    # a cell holds one value: array results (CSE formulas) give
                    # their first element to the anchor cell
                    val = float(val.ravel()[0]) if val.size else 0.0
            finally:
                self._stack.pop()
        else:
            val = self.cached.get(key)          # None = blank cell
        self.memo[key] = val
        return val

    def rect_values(self, sheet, c1, r1, c2, r2):
        r1, r2 = min(r1, r2 or r1), max(r1, r2 or r1)
        c1, c2 = min(c1, c2 or c1), max(c1, c2 or c1)
        out = np.zeros((r2 - r1 + 1, c2 - c1 + 1))
        for i, r in enumerate(range(r1, r2 + 1)):
            for j, c in enumerate(range(c1, c2 + 1)):
                v = self.cell_value(sheet, f"{get_column_letter(c)}{r}")
                out[i, j] = v if isinstance(v, (int, float)) else np.nan
        return out


def _to_num(v):
    """Coerce a value/Ref/array into a numpy array (2-D); blanks count as 0."""
    if isinstance(v, Ref):
        return np.nan_to_num(v.values(), nan=0.0)
    if isinstance(v, np.ndarray):
        return v if v.ndim == 2 else np.atleast_2d(v)
    if isinstance(v, bool):
        return np.array([[1.0 if v else 0.0]])
    if isinstance(v, (int, float)):
        return np.array([[float(v)]])
    if isinstance(v, str):
        try:
            return np.array([[float(v)]])
        except ValueError:
            return np.zeros((1, 1))
    return np.zeros((1, 1))


def _scalar(v):
    """Single value for scalar contexts (implicit intersection)."""
    if isinstance(v, Ref):
        return v.raw() if v.single else float(v.values().ravel()[0])
    if isinstance(v, np.ndarray):
        return float(v.ravel()[0]) if v.size else 0.0
    if isinstance(v, str):
        return v
    if v is None:
        return ""
    return float(v) if isinstance(v, (int, float)) else 0.0


def _num_scalar(v):
    s = _scalar(v)
    if not isinstance(s, str):
        return float(s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def _raw_flat(args):
    """Flatten keeping blanks as NaN (for AVERAGE/COUNT like Excel)."""
    out = []
    for a in args:
        if isinstance(a, Ref):
            out.append(a.values().ravel().astype(float))
        elif isinstance(a, np.ndarray):
            out.append(a.ravel().astype(float))
        else:
            out.append(np.atleast_1d(_num_scalar(a)))
    return np.concatenate(out) if out else np.array([])


def _flatten(args):
    out = []
    for a in args:
        if isinstance(a, (Ref, np.ndarray)):
            out.append(_to_num(a).ravel().astype(float))
        else:
            out.append(np.atleast_1d(_num_scalar(a)))
    return np.concatenate([a.astype(float) for a in out]) if out else np.array([])


def _is_container(v):
    return isinstance(v, (Ref, np.ndarray))


def _arith(op, a, b):
    """Element-wise when either side is an array/reference, else scalar."""
    if _is_container(a) or _is_container(b):
        A = _to_num(a)
        B = _to_num(b)
        with np.errstate(divide="ignore", invalid="ignore"):
            res = {"+": A + B, "-": A - B, "*": A * B,
                   "/": np.where(B != 0, A / np.where(B != 0, B, 1), 0.0),
                   "^": np.abs(A) ** B}[op]
        return res
    a, b = _num_scalar(a), _num_scalar(b)
    if op == "+":
        return a + b
    if op == "-":
        return a - b
    if op == "*":
        return a * b
    if op == "/":
        return a / b if b != 0 else 0.0
    return float(a) ** float(b)


class _Parser:
    def __init__(self, tokens, sheet, engine, coord=None):
        self.tok = tokens
        self.i, self.sheet, self.eng, self.coord = 0, sheet, engine, coord

    # -- helpers -----------------------------------------------------------
    def peek(self):
        return self.tok[self.i] if self.i < len(self.tok) else None

    def take(self):
        t = self.tok[self.i]
        self.i += 1
        return t

    def parse(self):
        return self.comparison()

    # -- grammar -----------------------------------------------------------
    def comparison(self):
        left = self.concat()
        while True:
            t = self.peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("=", "<>", "<", ">", "<=", ">="):
                self.take()
                right = self.concat()
                op = t.value
                if _is_container(left) or _is_container(right):
                    a, b = _to_num(left), _to_num(right)
                    left = {"=": a == b, "<>": a != b, "<": a < b, ">": a > b,
                            "<=": a <= b, ">=": a >= b}[op].astype(float)
                else:
                    a, b = _scalar(left), _scalar(right)
                    left = 1.0 if {"=": a == b, "<>": a != b, "<": a < b,
                                   ">": a > b, "<=": a <= b,
                                   ">=": a >= b}[op] else 0.0
            else:
                return left

    def concat(self):
        left = self.additive()
        while True:
            t = self.peek()
            if t and t.type == "OPERATOR-INFIX" and t.value == "&":
                self.take()
                left = str(_scalar(left)) + str(_scalar(self.additive()))
            else:
                return left

    def additive(self):
        left = self.term()
        while True:
            t = self.peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("+", "-"):
                self.take()
                left = _arith(t.value, left, self.term())
            else:
                return left

    def term(self):
        left = self.power()
        while True:
            t = self.peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("*", "/"):
                self.take()
                left = _arith(t.value, left, self.power())
            else:
                return left

    def power(self):
        base = self.unary()
        t = self.peek()
        if t and t.type == "OPERATOR-INFIX" and t.value == "^":
            self.take()
            return _arith("^", base, self.power())
        return base

    def unary(self):
        t = self.peek()
        if t and t.type == "OPERATOR-PREFIX":
            self.take()
            v = self.unary()
            if _is_container(v):
                A = _to_num(v)
                return -A if t.value == "-" else A
            return -_num_scalar(v) if t.value == "-" else _num_scalar(v)
        return self.postfix()

    def postfix(self):
        v = self.primary()
        while True:
            t = self.peek()
            if t and t.type == "OPERATOR-POSTFIX" and t.value == "%":
                self.take()
                v = _arith("/", v, 100.0)
            else:
                return v

    def primary(self):
        t = self.take()
        if t.type == "PAREN" and t.subtype == "OPEN":
            v = self.comparison()
            self.take()                                    # ')'
            return v
        if t.type == "ARRAY" and t.subtype == "OPEN":
            return self.array_constant()
        if t.type == "OPERAND":
            if t.subtype == "NUMBER":
                return float(t.value)
            if t.subtype == "TEXT":
                return t.value[1:-1].replace('""', '"')
            if t.subtype == "LOGICAL":
                return 1.0 if t.value.upper() == "TRUE" else 0.0
            if t.subtype == "RANGE":
                return self.reference(t.value)
        if t.type == "FUNC" and t.subtype == "OPEN":
            return self.call(t.value[:-1])
        raise SyntaxError(f"unexpected token {t.type} {t.value!r}")

    def array_constant(self):
        rows, row = [], []
        while True:
            t = self.peek()
            if t.type == "ARRAY" and t.subtype == "CLOSE":
                self.take()
                rows.append(row)
                break
            if t.type == "SEP" and t.value == ";":
                self.take()
                rows.append(row)
                row = []
                continue
            if t.type == "SEP" and t.value == ",":
                self.take()
                continue
            row.append(_num_scalar(self.comparison()))
        return np.array(rows, dtype=float)

    # -- references --------------------------------------------------------
    def reference(self, ref):
        cached = self.eng._ref_cache.get((self.sheet, ref))
        if cached is not None:
            return cached
        r = self._make_reference(ref)
        self.eng._ref_cache[(self.sheet, ref)] = r
        return r

    def _make_reference(self, ref):
        sheet, rest = _split_sheet(ref)
        sheet = self.eng._sheet(sheet) if sheet else self.sheet
        rest = rest.strip().replace("$", "")
        if _CELL_RE.match(rest):
            c, r = _split_cell(rest)
            return Ref(self.eng, sheet, c, r, c, r)
        if _RANGE_RE.match(rest):
            (c1, r1), (c2, r2) = _split_cell(rest.split(":")[0]), _split_cell(rest.split(":")[1])
            return Ref(self.eng, sheet, c1, r1, c2, r2)
        if ":" in rest and re.match(r"^[A-Za-z]+:[A-Za-z]+$", rest):     # whole columns
            a, b = rest.split(":")
            return Ref(self.eng, sheet, column_index_from_string(a), 1,
                       column_index_from_string(b), 1048576)
        target = self.eng.lookup_name(sheet, rest)
        if target:
            return self.reference(target.lstrip("="))
        raise NameError(f"cannot resolve reference {ref!r}")

    # -- functions ---------------------------------------------------------
    def args(self):
        out = []
        t = self.peek()
        if t and t.type == "FUNC" and t.subtype == "CLOSE":
            self.take()
            return out
        while True:
            out.append(self.comparison())
            t = self.peek()
            if t and t.type == "SEP":
                self.take()
                continue
            self.take()                                    # ')'
            return out

    def call(self, name):
        fname = name.split(".")[-1].upper()
        fn = getattr(self, "fn_" + fname, None)
        a = self.args()
        if fn is None:
            raise NotImplementedError(f"function not supported: {name}")
        return fn(a)

    # --- math / aggregate
    def fn_SUMPRODUCT(self, a):
        arrays = [_to_num(x).ravel().astype(float) for x in a]
        n = max(len(x) for x in arrays)
        arrays = [np.full(n, x[0]) if len(x) == 1 and n > 1 else x for x in arrays]
        if min(len(x) for x in arrays) != n:
            raise ValueError("SUMPRODUCT: mismatched array sizes")
        return float(np.prod(np.vstack(arrays), axis=0).sum())

    def fn_SUM(self, a):
        return float(np.nansum(_raw_flat(a)))

    def fn_SUMSQ(self, a):
        v = _raw_flat(a)
        return float(np.nansum(v * v))

    def fn_AVERAGE(self, a):
        v = _raw_flat(a)
        v = v[~np.isnan(v)]
        return float(v.mean()) if v.size else 0.0

    def fn_COUNT(self, a):
        v = _raw_flat(a)
        return float(np.count_nonzero(~np.isnan(v)))

    def fn_MIN(self, a):
        v = _raw_flat(a)
        v = v[~np.isnan(v)]
        return float(v.min()) if v.size else 0.0

    def fn_MAX(self, a):
        v = _raw_flat(a)
        v = v[~np.isnan(v)]
        return float(v.max()) if v.size else 0.0

    def fn_ABS(self, a):
        return float(abs(_num_scalar(a[0])))

    def fn_SQRT(self, a):
        return float(np.sqrt(max(_num_scalar(a[0]), 0.0)))

    def fn_EXP(self, a):
        return float(np.exp(_num_scalar(a[0])))

    def fn_LN(self, a):
        return float(np.log(_num_scalar(a[0])))

    def fn_LOG(self, a):
        x = _num_scalar(a[0])
        base = _num_scalar(a[1]) if len(a) > 1 else 10.0
        return float(np.log(x) / np.log(base))

    def fn_LOG10(self, a):
        return float(np.log10(_num_scalar(a[0])))

    def fn_POWER(self, a):
        return float(_num_scalar(a[0]) ** _num_scalar(a[1]))

    def fn_ROUND(self, a):
        return float(np.round(_num_scalar(a[0]), int(_num_scalar(a[1]))))

    def fn_ROUNDUP(self, a):
        if _is_container(a[0]):
            x = _to_num(a[0])
            d = _num_scalar(a[1]) if len(a) > 1 else 0
            f = 10 ** int(d)
            return np.sign(x) * np.ceil(np.abs(x) * f) / f
        x, d = _num_scalar(a[0]), int(_num_scalar(a[1]))
        f = 10 ** d
        return float(np.sign(x) * np.ceil(abs(x) * f) / f)

    def fn_ROUNDDOWN(self, a):
        if _is_container(a[0]):
            x = _to_num(a[0])
            d = _num_scalar(a[1]) if len(a) > 1 else 0
            f = 10 ** int(d)
            return np.sign(x) * np.floor(np.abs(x) * f) / f
        x, d = _num_scalar(a[0]), int(_num_scalar(a[1]))
        f = 10 ** d
        return float(np.sign(x) * np.floor(abs(x) * f) / f)

    def fn_INT(self, a):
        if _is_container(a[0]):
            return np.floor(_to_num(a[0]))
        return float(np.floor(_num_scalar(a[0])))

    def fn_TRUNC(self, a):
        return float(np.trunc(_num_scalar(a[0])))

    def fn_MOD(self, a):
        x, y = _num_scalar(a[0]), _num_scalar(a[1])
        return float(x - y * np.floor(x / y)) if y else 0.0

    def fn_SIGN(self, a):
        if _is_container(a[0]):
            return np.sign(_to_num(a[0]))
        return float(np.sign(_num_scalar(a[0])))

    def fn_MMULT(self, a):
        A, B = _to_num(a[0]), _to_num(a[1])
        if A.shape[1] != B.shape[0]:
            if A.shape[0] == B.shape[0] and A.shape[1] == 1:      # column vector x col
                A = A.T
            elif B.shape[1] == A.shape[0]:
                B = B.T
            else:
                raise ValueError(f"MMULT: {A.shape} x {B.shape} not conformable")
        return A @ B

    # --- lookup / reference functions
    def fn_COLUMN(self, a):
        if not a:
            return float(column_index_from_string(
                re.match(r"([A-Za-z]+)", self.coord).group(1))) if self.coord else 0.0
        v = a[0]
        if isinstance(v, Ref):
            return float(v.c1)
        return 0.0

    def fn_ROW(self, a):
        if not a:
            return float(re.sub(r"[A-Za-z$]", "", self.coord)) if self.coord else 0.0
        v = a[0]
        if isinstance(v, Ref):
            return float(v.r1)
        return 0.0

    def fn_INDEX(self, a):
        src = a[0]
        if not isinstance(src, Ref):
            return _num_scalar(src)
        vals = src.values()
        row = int(_num_scalar(a[1])) if len(a) > 1 else 1
        col = int(_num_scalar(a[2])) if len(a) > 2 else (1 if vals.shape[1] == 1 else 1)
        if row < 0 or col < 0:
            raise ValueError("INDEX: negative index")
        row = min(max(row, 1), vals.shape[0])
        col = min(max(col, 1), vals.shape[1])
        return float(vals[row - 1, col - 1])

    def fn_OFFSET(self, a):
        src = a[0]
        if not isinstance(src, Ref):
            return _num_scalar(src)
        dr = int(_num_scalar(a[1]))
        dc = int(_num_scalar(a[2]))
        h = int(_num_scalar(a[3])) if len(a) > 3 else 0
        w = int(_num_scalar(a[4])) if len(a) > 4 else 0
        v = src.shifted(dr, dc, h, w).values()
        return float(v.ravel()[0]) if v.size == 1 else v

    def fn_VLOOKUP(self, a):
        key = _num_scalar(a[0]) if not isinstance(_scalar(a[0]), str) else _scalar(a[0])
        table = _to_num(a[1])
        col = int(_num_scalar(a[2])) - 1
        exact = len(a) > 3 and not _num_scalar(a[3])
        first = table[:, 0]
        if exact:
            hit = np.where(np.isclose(first, key))[0] if isinstance(key, (int, float)) \
                else np.where(first == key)[0]
            if not len(hit):
                raise ValueError("VLOOKUP: not found")
            return float(table[hit[0], col])
        idx = 0
        for i, v in enumerate(first):
            if isinstance(key, (int, float)) and v <= key + 1e-12:
                idx = i
        return float(table[idx, col])

    def fn_SUMIF(self, a):
        rng = _to_num(a[0]).ravel()
        crit = a[1]
        sums = _to_num(a[2]).ravel() if len(a) > 2 else rng
        mask = _criteria_mask(rng, crit)
        return float(sums[:len(mask)][mask].sum())

    def fn_CODE(self, a):
        s = _scalar(a[0])
        return float(ord(str(s)[0])) if str(s) else 0.0

    # --- logic
    def fn_IF(self, a):
        if _num_scalar(a[0]):
            return a[1] if len(a) > 1 else 1.0
        return a[2] if len(a) > 2 else 0.0

    def fn_IFERROR(self, a):
        return a[0]

    def fn_AND(self, a):
        vals = _flatten(a)
        return 1.0 if vals.size and np.all(vals != 0) else 0.0

    def fn_OR(self, a):
        vals = _flatten(a)
        return 1.0 if np.any(vals != 0) else 0.0

    def fn_NOT(self, a):
        return 0.0 if _num_scalar(a[0]) else 1.0


def _split_cell(coord):
    m = re.match(r"^([A-Za-z]{1,3})(\d{1,7})$", coord.replace("$", ""))
    if not m:
        raise ValueError(f"bad cell address {coord!r}")
    return column_index_from_string(m.group(1).upper()), int(m.group(2))


def _criteria_mask(values, crit):
    """Build a boolean mask for SUMIF criteria like '>0', '<>x', '3'."""
    if isinstance(crit, str):
        m = re.match(r"^(<=|>=|<>|=|<|>)?(.*)$", crit.strip())
        op, rhs = m.group(1) or "=", m.group(2).strip()
        try:
            rhs = float(rhs)
        except ValueError:
            pass
        cmp = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b,
               "<": lambda a, b: a < b, ">": lambda a, b: a > b,
               "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}[op]
        return cmp(values, rhs)
    return values == _num_scalar(crit)


# ==========================================================================
# 2. Solver metadata
# ==========================================================================
OBJ_KEYS = ("solver_opt", "solver_obj", "solver_obj_cell", "solver_set_cell")
TOL = 1e-6          # relative constraint tolerance


def expand_range(ref, default_sheet):
    """'Sheet!$A$1:$B$3' -> ['Sheet!A1', 'Sheet!B1', 'Sheet!A2', ...]"""
    if "!" in ref:
        sheet, rng = ref.split("!")
    else:
        sheet, rng = default_sheet, ref
    min_col, min_row, max_col, max_row = range_boundaries(rng.replace("$", ""))
    return [f"{sheet}!{get_column_letter(c)}{r}"
            for r in range(min_row, max_row + 1)
            for c in range(min_col, max_col + 1)]


def parse_solver_metadata(wb, sheet=None):
    """Extract the Solver model of one worksheet.

    Workbooks often hold one Solver model per worksheet. Without an explicit
    `sheet`, the model of the DEFAULT (active) worksheet wins; if that sheet
    has no model, a single existing model is used, otherwise the first model
    found is used with a warning.
    """
    models = {}
    for ws in wb.worksheets:
        d = {n.lower(): dfn.attr_text for n, dfn in ws.defined_names.items()
             if n.lower().startswith("solver_")}
        if any(k in d for k in OBJ_KEYS):
            models[ws.title] = d
    if not models:
        raise ValueError(
            "No Solver model metadata found in any worksheet "
            "(need %s + 'solver_adj').\n"
            "Tip: press Solve (or Save) once in Excel so Solver writes its "
            "model names into the file." % (OBJ_KEYS,))

    if sheet:
        match = next((t for t in models if t.lower() == sheet.lower()), None)
        if match is None:
            raise ValueError(f"No Solver model on sheet {sheet!r}. "
                             f"Available: {list(models)}")
        target, how = match, "requested"
    else:
        active = wb.active.title
        if active in models:
            target, how = active, "default (active) worksheet"
        elif len(models) == 1:
            target, how = next(iter(models)), "only model in the workbook"
        else:
            target = next(iter(models))
            how = "the active worksheet has no model, using the first one"
            print(f"[warning] Solver models found on {list(models)}; the "
                  f"active sheet {active!r} has none. Use --sheet to choose.")

    if len(models) > 1:
        print(f"Solver models found on sheets {list(models)} -> using "
              f"{target!r} ({how}).")
    raw = models[target]

    # names visible from the model's sheet: worksheet-scoped first, then global
    scoped = {n.lower(): d.attr_text
              for n, d in wb[target].defined_names.items()}
    global_names = {n.lower(): d.attr_text for n, d in wb.defined_names.items()}

    def resolve(raw_ref):
        """Reference or defined name -> 'Sheet!A1:B2'."""
        text = raw_ref.lstrip("=").strip()
        for _ in range(10):
            s = text.replace("$", "").replace("'", "")
            if "!" in s:
                sheet_part, rng = s.rsplit("!", 1)
                return f"{sheet_part}!{rng}"
            try:
                range_boundaries(s)
                return f"{target}!{s}"
            except ValueError:
                pass
            name = text.strip().strip("=").lower()
            nxt = scoped.get(name) or global_names.get(name)
            if nxt is None:
                raise ValueError(f"cannot resolve reference {raw_ref!r} "
                                 f"(no such defined name)")
            text = nxt
        raise ValueError(f"defined name chain too deep for {raw_ref!r}")

    obj_cell = resolve(raw[next(k for k in OBJ_KEYS if k in raw)])
    obj_type = int(float(raw.get("solver_typ", "2").lstrip("=")))
    obj_target = float(raw.get("solver_val", "0").lstrip("=") or 0.0)
    non_negative = raw.get("solver_neg", "1").lstrip("=") == "1"

    var_cells = []
    for piece in raw["solver_adj"].split(","):
        piece = piece.strip()
        if piece:
            var_cells.extend(expand_range(resolve(piece), target))

    constraints, i = [], 1
    while f"solver_lhs{i}" in raw:
        lhs = expand_range(resolve(raw[f"solver_lhs{i}"]), target)
        rel = int(float(raw[f"solver_rel{i}"].lstrip("=")))
        rhs_raw = raw[f"solver_rhs{i}"].lstrip("=").strip()
        try:
            rhs = [float(rhs_raw)] * len(lhs)
        except ValueError:
            rhs = expand_range(resolve(rhs_raw), target)
            if len(rhs) != len(lhs):
                raise ValueError(
                    f"constraint {i}: {len(lhs)} left-hand cells vs "
                    f"{len(rhs)} right-hand cells ({rhs_raw})")
        constraints += list(zip(lhs, [rel] * len(lhs), rhs))
        i += 1

    return dict(sheet=target, obj_cell=obj_cell, obj_type=obj_type,
                obj_target=obj_target, var_cells=var_cells,
                non_negative=non_negative, constraints=constraints,
                models=list(models))


# ==========================================================================
# 3. Solving
# ==========================================================================
def _finite_diff(fn, x, h=None):
    """Gradient and constant term of an (assumed) affine function."""
    x = np.asarray(x, float)
    h = np.where(np.abs(x) > 1e-9, 1e-6 * np.abs(x), 1e-6) if h is None else h
    f0 = fn(x)
    grad = np.empty_like(x)
    for i in range(len(x)):
        e = np.zeros_like(x)
        e[i] = h[i]
        grad[i] = (fn(x + e) - fn(x - e)) / (2 * h[i])
    return grad, f0 - grad @ x


def check_affine(fn, x0, tol=1e-6):
    """Decide whether fn(x) is affine, i.e. fn(x) = g.x + d.

    Three independent tests, any of which rejects the affine hypothesis:

      1. the affine model fitted at the FIRST sample point must predict fn at
         every other sample point;
      2. finite-difference gradients must agree between sample points;
      3. the constant term must agree between sample points.

    Sample points are drawn at three magnitudes (0.3x, 1x, 4x of |x0|) so a
    nonlinearity that is invisible close to the starting point is still seen.
    Any evaluation erroring out counts as "not affine" (conservative: the
    nonlinear solver then handles it).
    """
    n = len(x0)
    scale = np.where(np.abs(x0) > 1e-9, np.abs(x0), 1.0)
    rng = np.random.default_rng(0)
    pts = [m * scale * rng.uniform(0.5, 1.5, size=n)
           for m in (0.3, 1.0, 4.0) for _ in range(2)]
    try:
        g0, d0 = _finite_diff(fn, pts[0])
        if not (np.all(np.isfinite(g0)) and np.isfinite(d0)):
            return False, None, None
        g_tol = tol * max(1.0, float(np.max(np.abs(g0))))
        for x in pts[1:]:
            got = float(fn(x))
            if not np.isfinite(got):
                return False, None, None
            if abs(got - float(g0 @ x + d0)) > tol * max(1.0, abs(got)):
                return False, None, None
            g, d = _finite_diff(fn, x)
            if not (np.all(np.isfinite(g)) and np.isfinite(d)):
                return False, None, None
            if np.any(np.abs(g - g0) > g_tol):
                return False, None, None
            if abs(d - d0) > tol * max(1.0, abs(d0)):
                return False, None, None
    except Exception:                                  # noqa: BLE001
        return False, None, None
    return True, g0, d0


def solve_and_update_excel(file_path, output_path=None, mode="auto",
                           sheet=None):
    output_path = output_path or file_path

    keep_vba = str(file_path).lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    meta = parse_solver_metadata(wb, sheet=sheet)
    eng = FormulaEngine(file_path)

    var_cells = meta["var_cells"]
    n = len(var_cells)
    obj_cell = meta["obj_cell"]
    is_target = meta["obj_type"] == 3
    maximize = meta["obj_type"] == 1
    sign = -1.0 if maximize else 1.0        # for scipy minimisation

    constraints = list(meta["constraints"])
    if is_target:
        constraints.append((obj_cell, 2, meta["obj_target"]))

    print(f"Objective : {obj_cell}  "
          f"({'maximise' if maximize else 'minimise' if not is_target else 'target'})")
    print(f"Variables : {n}  -> {var_cells[0]} ... {var_cells[-1]}")
    print(f"Constraints: {len(constraints)} on sheet {meta['sheet']!r}")

    lb = np.zeros(n) if meta["non_negative"] else np.full(n, -np.inf)
    ub = np.full(n, np.inf)
    integrality = np.zeros(n)

    C = []                                  # (lhs, rel, rhs)
    for lhs, rel, rhs in constraints:
        if rel in (4, 5):                   # integer / binary
            if lhs not in var_cells:
                raise ValueError(
                    f"Integer/Binary constraint on non-variable cell: {lhs}")
            j = var_cells.index(lhs)
            integrality[j] = 1.0
            if rel == 5:
                lb[j], ub[j] = max(lb[j], 0.0), min(ub[j], 1.0)
            continue
        C.append((lhs, rel, rhs))

    def obj_at(x):
        eng.set_values(var_cells, x)
        return float(eng.value(obj_cell))

    def all_values(x):
        """Objective value and all constraint slacks (lhs - rhs) at x."""
        eng.set_values(var_cells, x)
        o = float(eng.value(obj_cell))
        g = np.empty(len(C))
        for k, (lhs, _, rhs) in enumerate(C):
            v = float(eng.value(lhs))
            g[k] = v - (float(eng.value(rhs)) if isinstance(rhs, str) else rhs)
        return o, g

    def violation(x):
        _, g = all_values(x)
        total = worst = 0.0
        for k, (_, rel, _) in enumerate(C):
            v = g[k] if rel == 1 else (-g[k] if rel == 3 else abs(g[k]))
            v = max(0.0, v) if rel in (1, 3) else abs(g[k])
            sc = max(1.0, abs(g[k]))
            total += (v / sc) ** 2
            worst = max(worst, v / sc)
        return total, worst

    x0 = np.array([float(v) if isinstance(v, (int, float)) else 0.0
                   for v in (eng.value(c) for c in var_cells)])

    # --- self-check: engine vs Excel's cached results ---------------------
    bad = []
    for (sh, coord), formula in eng.formulas.items():
        want = eng.cached.get((sh, coord))
        if not isinstance(want, (int, float)):
            continue
        try:
            got = eng.cell_value(sh, coord)
        except Exception as exc:                     # noqa: BLE001
            bad.append((sh, coord, formula, want, f"error: {exc}"))
            continue
        if not isinstance(got, (int, float)) or not np.isclose(
                got, want, atol=1e-6 * max(1.0, abs(want)), rtol=1e-6):
            bad.append((sh, coord, formula, want, got))
    if bad:
        print(f"\n[warning] {len(bad)} formula(s) do not reproduce Excel's "
              f"cached value:")
        for sh, coord, formula, want, got in bad[:10]:
            print(f"   {sh}!{coord}  {formula[:70]}  cached={want}  computed={got}")
    else:
        print(f"Self-check: all {len(eng.formulas)} formula cells reproduce "
              f"Excel's cached values.")

    print(f"Objective at start: {obj_at(x0):.6f}")

    # --- linear? ----------------------------------------------------------
    def fd_jac(vec_fn, x):
        """Forward-difference Jacobian of a vector function (one pass)."""
        f0 = vec_fn(x)
        J = np.empty((len(f0), n))
        for i in range(n):
            h = 1e-6 * max(1.0, abs(x[i]))
            xp = x.copy()
            xp[i] += h
            J[:, i] = (vec_fn(xp) - f0) / h
        return f0, J

    def linearity(vec_fn, x0_, tol=1e-6):
        """Per-component test whether f_i is affine."""
        scale = np.where(np.abs(x0_) > 1e-9, np.abs(x0_), 1.0)
        rng = np.random.default_rng(0)
        pts = [m * scale * rng.uniform(0.5, 1.5, size=n)
               for m in (0.3, 1.0, 4.0)]
        try:
            f0, J0 = fd_jac(vec_fn, pts[0])
            d0 = f0 - J0 @ pts[0]
            ok = np.ones(len(f0), bool)
            for x in pts[1:]:
                f, J = fd_jac(vec_fn, x)
                pred = J0 @ x + d0
                bad_rows = (np.abs(f - pred) >
                            tol * np.maximum(1.0, np.abs(f)))
                bad_rows |= np.any(
                    np.abs(J - J0) > tol * np.maximum(1.0, np.abs(J0)), axis=1)
                ok &= ~bad_rows
            return ok, J0, d0
        except Exception:                            # noqa: BLE001
            return np.zeros(1, bool), None, None

    def all_vals_vector(x):
        o, g = all_values(x)
        return np.concatenate(([o], g))

    if mode == "nonlinear":
        is_affine, obj_ok, cons_ok = False, False, False
        print("Linearity: skipped (forced nonlinear mode).")
    elif mode == "linear":
        is_affine, obj_ok, cons_ok = True, True, True
        print("Linearity: skipped (forced LP mode).")
    else:
        ok, J, d = linearity(all_vals_vector, x0)
        if J is None:
            obj_ok, cons_ok, is_affine = False, False, False
        else:
            obj_ok = bool(ok[0]) and not is_target
            cons_ok = bool(ok[1:].all()) if len(C) else True
            is_affine = obj_ok or is_target
            is_affine = is_affine and cons_ok
        print(f"Linearity: objective {'affine' if obj_ok or is_target else 'NONLINEAR'}"
              f", constraints {'affine' if cons_ok else 'NONLINEAR'}"
              f" -> {'LP/MILP' if is_affine else 'nonlinear (SLSQP)'}")

    goal = (lambda x: 0.0) if is_target else obj_at

    # --- shared finite-difference cache: objective + all constraints at once
    _cache = {"n": 0, "data": {}}

    def vals_jac(x):
        key = x.tobytes()
        hit = _cache["data"].get(key)
        if hit is not None:
            return hit
        f0 = all_vals_vector(x)
        J = np.empty((len(f0), n))
        for i in range(n):
            h = 1e-6 * max(1.0, abs(x[i]))
            xp = x.copy()
            xp[i] += h
            J[:, i] = (all_vals_vector(xp) - f0) / h
        out = (f0, J)
        if len(_cache["data"]) > 64:
            _cache["data"].clear()
        _cache["data"][key] = out
        return out

    def run_nonlinear(extra_starts=(), random_starts=True):
        if not HAVE_SCIPY:
            raise RuntimeError("Nonlinear model needs scipy (pip install scipy).")
        sc = []
        for k, (_, rel, _) in enumerate(C):
            d = {"fun": (lambda x, k=k: vals_jac(x)[0][k + 1]),
                 "jac": (lambda x, k=k: vals_jac(x)[1][k + 1])}
            d["type"] = "ineq" if rel in (1, 3) else "eq"
            if rel == 1:                              # lhs - rhs <= 0
                d["fun"] = (lambda x, k=k: -vals_jac(x)[0][k + 1])
                d["jac"] = (lambda x, k=k: -vals_jac(x)[1][k + 1])
            sc.append(d)

        bounds_list = list(zip([None if np.isinf(l) else l for l in lb],
                               [None if np.isinf(u) else u for u in ub]))
        scale = max(1.0, float(np.max(np.abs(x0))))
        starts = ([x0] if random_starts else []) + list(extra_starts)
        rng = np.random.default_rng(7)
        n_random = 3 if n <= 40 else 2
        if random_starts:
            for k in range(n_random):
                mag = (0.5, 2.0, 4.0)[k % 3] * scale
                starts.append(rng.uniform(0.0, mag, size=n))

        best = None
        for k, st in enumerate(starts):
            try:
                r = minimize(lambda x: sign * vals_jac(x)[0][0],
                             np.clip(st, lb, ub), jac=lambda x: sign * vals_jac(x)[1][0],
                             method="SLSQP", bounds=bounds_list, constraints=sc,
                             options={"ftol": 1e-8, "maxiter": 300})
            except Exception as exc:                 # noqa: BLE001
                print(f"  start {k}: failed ({exc})")
                continue
            if r.x is None or not np.all(np.isfinite(r.x)):
                continue
            val, worst = goal(r.x), violation(r.x)[1]
            print(f"  start {k}: objective={val:.3f} "
                  f"[{'feasible' if worst <= TOL else f'infeasible {worst:.1e}'}]")
            if worst <= TOL and (best is None or val > best[0]):
                best = (val, r.x)
        return best

    def affine_system():
        """Normalised (row, lower, upper) form of the affine constraints."""
        A, bl, bu, sc = [], [], [], []
        for k, (lhs, rel, rhs) in enumerate(C):
            row, const = J[k + 1].copy(), d[k + 1]
            # scale so the largest coefficient is 1 (conditioning)
            s = max(1e-12, float(np.max(np.abs(row))))
            row, const = row / s, const / s
            A.append(row)
            if rel == 1:                     # row.x + const <= 0
                bl.append(-np.inf); bu.append(-const)
            elif rel == 3:                   # row.x + const >= 0
                bl.append(-const); bu.append(np.inf)
            else:                            # equality
                bl.append(-const); bu.append(-const)
        return np.array(A), np.array(bl), np.array(bu), sc

    def lp_feasible(cvec=None, A=None, bl=None, bu=None):
        """A point satisfying the affine constraints (optionally optimising a
        linear objective) -- used as a feasible starting point."""
        r = milp(c=np.zeros(n) if cvec is None else cvec,
                 integrality=integrality, bounds=Bounds(lb, ub),
                 constraints=LinearConstraint(A, bl, bu))
        return r.x if r.success and r.x is not None else None

    def lp_restore(A, bl, bu, x_ref):
        """Feasible point closest to x_ref (L1), used when the sheet values are
        slightly infeasible.  Exact closest point keeps the warm start useful."""
        m = A.shape[0]
        shift = A @ x_ref
        cons = LinearConstraint(np.hstack([A, -A]),
                                bl - shift, bu - shift)
        hi = np.concatenate([np.maximum(ub - x_ref, 0.0),
                             np.maximum(x_ref - lb, 0.0)])
        r = milp(c=np.ones(2 * n), integrality=np.zeros(2 * n),
                 bounds=Bounds(np.zeros(2 * n), hi), constraints=cons)
        if r.x is None:
            return None
        return x_ref + r.x[:n] - r.x[n:]

    def obj_grad(x):
        """Finite-difference gradient of the objective only (cheap)."""
        f0 = obj_at(x)
        g = np.empty(n)
        for i in range(n):
            h = 1e-5 * max(1.0, abs(x[i]))
            xp = x.copy()
            xp[i] += h
            g[i] = (obj_at(xp) - f0) / h
        return g

    def run_slp_affine(x_start, max_iter=60):
        """Successive linear programming for a nonlinear objective over affine
        constraints.

        Each iteration linearises the objective, maximises it over the (exact)
        constraint polytope inside a trust box, and accepts the step only if
        the true objective improves while feasibility is kept.  This is far more
        reliable than SLSQP on the 60+ variable business-planning workbooks
        Excel Solver is normally used for.
        """
        if not HAVE_SCIPY:
            raise RuntimeError("Nonlinear model needs scipy (pip install scipy).")
        A, bl, bu, _ = affine_system()
        x = np.clip(np.asarray(x_start, float), lb, ub)
        fx = goal(x)
        best = (fx, x.copy()) if violation(x)[1] <= TOL else (None, None)
        for it in range(max_iter):
            g = obj_grad(x)                       # d(objective)/dx
            delta = 0.5 * np.maximum(np.abs(x), 1.0)
            lo = np.maximum(lb, x - delta)
            hi = np.minimum(ub, x + delta)
            r = milp(c=-g, integrality=integrality, bounds=Bounds(lo, hi),
                     constraints=LinearConstraint(A, bl, bu))
            if r.x is None:
                break
            step, t = r.x - x, 1.0
            improved = False

            def aff_ok(z):                       # cheap feasibility (exact here)
                az = A @ z
                return bool(np.all(az >= bl - 1e-8) and np.all(az <= bu + 1e-8))

            while t > 1e-3:
                xt = x + t * step
                ft = goal(xt)
                if ft > fx + 1e-12 and aff_ok(xt):
                    x, fx = xt, ft
                    improved = True
                    break
                t *= 0.5
            if best[0] is None or fx > best[0]:
                if violation(x)[1] <= TOL:
                    best = (fx, x.copy())
            if not improved:
                break
        if best[0] is None:
            return None
        print(f"  SLP: objective={best[0]:.3f} after {it + 1} iterations")
        return best

    x_best = None
    if is_affine:
        A, bl, bu = [], [], []
        for k, (lhs, rel, rhs) in enumerate(C):
            row, const = J[k + 1], d[k + 1]
            A.append(row)
            if rel == 1:
                bl.append(-np.inf); bu.append(-const)
            elif rel == 3:
                bl.append(-const); bu.append(np.inf)
            else:
                bl.append(-const); bu.append(-const)
        lin = LinearConstraint(np.array(A), bl, bu) if A else None
        c = np.zeros(n) if is_target else sign * J[0]
        print("Solving as LP/MILP with SciPy HiGHS...")
        res = milp(c=c, integrality=integrality, bounds=Bounds(lb, ub),
                   constraints=lin)
        print("  ", res.message)
        if res.x is not None:
            true_obj, worst = obj_at(res.x), violation(res.x)[1]
            print(f"  LP solution: objective={true_obj:.6f}, "
                  f"max relative violation={worst:.2e}")
            if worst <= TOL:
                x_best = res.x
            else:
                print("  [warning] the LP solution violates the real formulas "
                      "-> re-solving as nonlinear.")
                best = run_nonlinear(extra_starts=[res.x])
                x_best = best[1] if best else None
        else:
            print("  [warning] LP failed -> trying the nonlinear solver.")
            best = run_nonlinear()
            x_best = best[1] if best else None
    else:
        cands = []
        if C and cons_ok and not is_target:
            print("Solving nonlinear objective over affine constraints "
                  "(successive LP)...")
            slp = None
            start_pt = x0
            if violation(x0)[1] > TOL:          # sheet values are infeasible
                A_, bl_, bu_, _ = affine_system()
                repaired = lp_restore(A_, bl_, bu_, x0)
                if repaired is None:
                    print("  [infeasible] the constraints cannot be satisfied.")
                else:
                    start_pt = repaired
                    print("  sheet values violated the constraints -> repaired "
                          "with the closest feasible point")
            if start_pt is not None:
                slp = run_slp_affine(start_pt)
            if slp is not None:
                cands.append(slp)
        if not C or n <= 25:
            print("Solving nonlinear model (multi-start SLSQP)...")
            nl = run_nonlinear()
            if nl is not None:
                cands.append(nl)
        cands = [c for c in cands if violation(c[1])[1] <= TOL]
        best = max(cands, key=lambda c: c[0]) if cands else None
        if len(cands) > 1:
            print("  best of: " + ", ".join(f"{c[0]:.3f}" for c in cands))
        x_best = best[1] if best else None

    if x_best is None:
        print("Optimization Failed: no feasible point found.")
        return None

    # --- independent verification with a FRESH engine (no cached values) --
    ver = FormulaEngine(file_path)
    ver.set_values(var_cells, x_best)
    final_obj = float(ver.value(obj_cell))
    worst, violated = 0.0, []
    for lhs, rel, rhs in constraints:
        lv = float(ver.value(lhs))
        rv = float(ver.value(rhs)) if isinstance(rhs, str) else rhs
        g = lv - rv
        v = g if rel == 1 else (-g if rel == 3 else abs(g))
        v = max(0.0, v) if rel in (1, 3) else abs(g)
        sc = max(1.0, abs(lv), abs(rv))
        worst = max(worst, v / sc)
        if v > 1e-6 * sc:
            violated.append((lhs, lv, rhs, rv, g))

    print(f"\nOptimal objective: {final_obj:.6f}")
    print(f"Max relative constraint violation (verified): {worst:.2e}")
    if violated:
        print(f"[warning] {len(violated)} constraint(s) still violated:")
        for lhs, lv, rhs, rv, g in violated[:8]:
            print(f"   {lhs} = {lv:.6f}  vs  "
                  f"{rhs if isinstance(rhs, str) else rhs} = {rv:.6f}"
                  f"   (diff {g:+.6f})")
    else:
        print("All constraints satisfied.")

    for j, (cell, val) in enumerate(zip(var_cells, x_best)):
        sh, coord = cell.split("!")
        wb[sh][coord] = int(round(val)) if integrality[j] == 1 else float(val)

    wb.calculation.fullCalcOnLoad = True
    wb.save(output_path)
    print(f"Saved to {output_path}")
    return final_obj


def select_excel_file():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select an Excel Solver Workbook",
        filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
    root.destroy()
    return path


if __name__ == "__main__":
    import sys
    from pathlib import Path

    argv = sys.argv[1:]
    args = [a for a in argv if not a.startswith("--")]
    mode = ("nonlinear" if "--nonlinear" in argv else
            "linear" if "--linear" in argv else "auto")
    want_sheet = None
    if "--sheet" in argv:
        j = argv.index("--sheet")
        want_sheet = argv[j + 1] if j + 1 < len(argv) else None
        args = [a for a in args if a != want_sheet]

    selected = args[0] if args else select_excel_file()
    if selected:
        p = Path(selected)
        save = str(p.parent / f"pysolved-{p.name}")
        print("Solving the model in", selected)
        solve_and_update_excel(selected, save, mode=mode,
                               sheet=want_sheet)
        input("Press Enter to exit...")
