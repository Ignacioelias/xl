"""Excel value semantics, as far as the builders need them.

A cell value read from the cache is one of: None (blank), float/int, bool, str, XlError.
The helpers here reproduce what Excel does with those in the handful of operations the
builders emit. Where Excel's behaviour depends on the locale (text that looks like a number
or a date being coerced in arithmetic), nothing is guessed: `Unmirrored` is raised and the
cell is reported, so a computed value is never a plausible-looking wrong number.

Not mirrored on purpose: Excel's "close to zero" adjustment of a formula's final + or -
(a residual of 1e-15 can be shown as exactly 0). Compare check rows with a tolerance.
"""
import re


class XlError:
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = code

    def __eq__(self, other):
        return isinstance(other, XlError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return self.code


NA = XlError("#N/A")
VALUE = XlError("#VALUE!")
DIV0 = XlError("#DIV/0!")
REF = XlError("#REF!")
ERROR_CODES = ("#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NUM!", "#NULL!",
               "#GETTING_DATA", "#SPILL!", "#CALC!")


class Unmirrored(Exception):
    """Excel's result here depends on something the mirror does not reproduce."""


def is_err(v):
    return isinstance(v, XlError)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# text Excel may coerce to a number in arithmetic or in a *IFS criterion: plain numbers,
# thousands separators, percentages, currency, dates and times. How it coerces them depends
# on the locale, so the mirror refuses instead of guessing.
_COERCIBLE = re.compile(
    r"^\s*[-+(]?\s*[$€£¥]?\s*[-+]?(\d[\d.,\s]*|[.,]\d+)([eE][-+]?\d+)?\s*%?\s*\)?\s*$"
    r"|^\s*\d{1,4}\s*[-/.]\s*\d{1,2}(\s*[-/.]\s*\d{1,4})?(\s+\d{1,2}:\d{2}(:\d{2})?)?\s*$"
    r"|^\s*\d{1,2}:\d{2}(:\d{2})?(\s*[AaPp][Mm])?\s*$")


def coercible_text(v):
    return isinstance(v, str) and (_COERCIBLE.match(v) is not None
                                   or v.strip().upper() in ("TRUE", "FALSE"))


def arith(v, where=""):
    """A cell value used as an operand of + - * / (or inside an array product)."""
    if v is None:
        return 0.0
    if isinstance(v, XlError):
        return v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        if v.strip() == "":
            return VALUE
        if coercible_text(v):
            raise Unmirrored(f"text {v!r} used in arithmetic{where}: Excel's coercion is locale-dependent")
        return VALUE
    raise Unmirrored(f"unsupported value {v!r}{where}")


def _first_err(*xs):
    for x in xs:
        if isinstance(x, XlError):
            return x
    return None


def add(a, b):
    return _first_err(a, b) or a + b


def sub(a, b):
    return _first_err(a, b) or a - b


def mul(a, b):
    return _first_err(a, b) or a * b


def div(a, b):
    e = _first_err(a, b)
    if e:
        return e
    if b == 0:
        return DIV0
    return a / b


def neg(a):
    return a if isinstance(a, XlError) else -a


def sum_refs(values):
    """SUM over cell references: text, booleans and blanks are ignored, errors propagate."""
    total = 0.0
    for v in values:
        if isinstance(v, XlError):
            return v
        if is_num(v):
            total += v
    return total


def crit_equal_num(v, n, where=""):
    """SUMIFS/COUNTIFS numeric criterion: a blank cell never matches (a formula returning 0
    does), a number matches when equal, text Excel could coerce is refused."""
    if v is None or isinstance(v, (bool, XlError)):
        return False
    if is_num(v):
        return float(v) == float(n)
    if coercible_text(v):
        raise Unmirrored(f"number-like text {v!r} in a criteria range{where}")
    return False


def match_exact_num(values, n):
    """MATCH(n, range, 0): 1-based position of the first number equal to n, or None.
    MATCH does not coerce types, so text ("2025", "FY25") never matches a number."""
    for i, v in enumerate(values, start=1):
        if is_num(v) and float(v) == float(n):
            return i
    return None


def _cmp_key(v):
    """Excel comparison ordering: numbers < text < logicals; blank compares as 0 (or "")."""
    if v is None:
        return (0, 0.0)
    if isinstance(v, bool):
        return (2, v)
    if is_num(v):
        return (0, float(v))
    return (1, str(v).lower())


def cmp(v, n, op):
    """`range <op> number` element: True/False, or the XlError of the cell. Comparisons do
    not coerce: any text sorts above every number, so a stray text year passes `>=` and
    fails `<=` (which is what the column-count check row exists to catch)."""
    if isinstance(v, XlError):
        return v
    a, b = _cmp_key(v), (0, float(n))
    return {"=": a == b, ">=": a >= b, "<=": a <= b, ">": a > b, "<": a < b}[op]


def cmp_eq_value(v, target):
    """`cell = target` where target is a number or text (IF tests). Text compares case-insensitively."""
    if isinstance(v, XlError):
        return v
    if isinstance(target, str):
        return isinstance(v, str) and v.lower() == target.lower() or (v is None and target == "")
    return cmp(v, target, "=")


def wildcard_match(pattern, text):
    """SUMIFS text criterion: case-insensitive, * and ? wildcards, ~ escapes."""
    out, i = [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern):
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.fullmatch("".join(out), text, re.I | re.S) is not None


def same_value(a, b, rel=1e-12, abs_tol=1e-12):
    """Cached value equality for verification (floats with a tight tolerance)."""
    if is_num(a) and is_num(b):
        return abs(a - b) <= max(abs_tol, rel * max(abs(a), abs(b)))
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b or a == b
    if a in (None, "") and b in (None, ""):
        return True
    return a == b


def to_json(v):
    if isinstance(v, XlError):
        return {"error": v.code}
    return v
