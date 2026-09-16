"""Formula builders. Each builder writes ONE Excel idiom and mirrors it in Python.

`formula(cx)` returns the formula text without the leading "=".
`value(cx)` returns what Excel would cache for that formula, from the model's cached values
(`cx.model`) and the block's own computed cells (`cx.key_value`, `cx.cell_value`).
`needs(cx)` yields the model ranges `value` reads, as (sheet, r0, r1, c0, c1).

Keeping the two halves in one class is the point: the injected value and the live formula
come from the same definition, so they cannot drift apart. Every year is written into the
formula as a LITERAL (or an absolute reference to a numeric input cell), never as a pointer
at a header label, so relabelling a header can never blank a block.
"""
import re
from dataclasses import dataclass

from . import a1
from . import xlsem as X


class SpecError(ValueError):
    pass


# ------------------------------------------------------------------------------ periods
@dataclass(frozen=True)
class Period:
    year: int
    ref: str = None          # absolute address of a numeric input cell, e.g. "$D$7"

    def text(self):
        return self.ref if self.ref else str(self.year)

    def shift(self, k):
        if not k:
            return self
        ref = None if self.ref is None else f"({self.ref}{'+' if k > 0 else '-'}{abs(k)})"
        return Period(self.year + k, ref)


_TOKEN = re.compile(r"^\$(\w+)\s*([-+]\s*\d+)?$")


def resolve_period(tok, periods, where):
    """tok: an int year, or "$year" / "$open" / "$close", optionally "+1" / "-1"."""
    if isinstance(tok, bool):
        raise SpecError(f"{where}: bad period {tok!r}")
    if isinstance(tok, int):
        return Period(tok)
    if isinstance(tok, str):
        m = _TOKEN.match(tok.strip())
        if m and m.group(1) in periods and periods[m.group(1)] is not None:
            k = int(m.group(2).replace(" ", "")) if m.group(2) else 0
            return periods[m.group(1)].shift(k)
    raise SpecError(f"{where}: period {tok!r} is not an int or one of "
                    f"{['$' + k for k, v in periods.items() if v is not None]}")


# --------------------------------------------------------------------------------- grids
@dataclass(frozen=True)
class Grid:
    """A model sheet's period axis: the row that holds a year per column, and the columns."""
    name: str
    sheet: str
    year_row: int
    c0: int
    c1: int

    def years(self):
        return a1.row_range(self.sheet, self.year_row, self.c0, self.c1)

    def row(self, r):
        return a1.row_range(self.sheet, r, self.c0, self.c1)

    def cols(self):
        return range(self.c0, self.c1 + 1)

    def need(self, r):
        return (self.sheet, r, r, self.c0, self.c1)


def _num(x):
    x = float(x)
    return str(int(x)) if x.is_integer() and abs(x) < 1e15 else repr(x)


def _wrap(expr, scale, sign, compound):
    s = f"({expr})" if compound and (scale != 1 or sign == -1) else expr
    if scale != 1:
        s = f"{s}/{_num(scale)}"
    if sign == -1:
        s = f"-{s}"
    return s


def _finish(v, scale, sign, raw_ok=True):
    """Value of _wrap(...): raw pass-through when nothing is applied (a formula that points at
    a blank cell caches 0), arithmetic otherwise."""
    if scale == 1 and sign == 1 and raw_ok:
        return 0.0 if v is None else v
    x = X.arith(v)
    if sign == -1:
        x = X.neg(x)
    if scale != 1:
        x = X.div(x, float(scale))
    return x


class Builder:
    fn = None
    raw = False              # True when the formula may return text (links, lookups)

    def __init__(self, p, ctx, where):
        self.p = p
        self.where = where
        self._grids = ctx["grids"]

    # --- parameter helpers
    def grid(self):
        name = self.p.get("grid")
        if name not in self._grids:
            raise SpecError(f"{self.where}: grid {name!r} not defined (grids: {sorted(self._grids)})")
        return self._grids[name]

    def rows(self):
        r = self.p.get("rows", self.p.get("row"))
        rows = [r] if isinstance(r, int) else r
        if not rows or not all(isinstance(x, int) and x > 0 for x in rows):
            raise SpecError(f"{self.where}: 'row' (int) or 'rows' (list of int) required")
        return list(rows)

    def scale_sign(self):
        scale = self.p.get("scale", 1)
        sign = self.p.get("sign", 1)
        if not isinstance(scale, (int, float)) or isinstance(scale, bool) or scale == 0:
            raise SpecError(f"{self.where}: scale must be a non-zero number")
        if sign not in (1, -1):
            raise SpecError(f"{self.where}: sign must be 1 or -1")
        return scale, sign

    # --- interface
    def needs(self, cx):
        return ()

    def deps(self):
        return ()

    def formula(self, cx):
        raise NotImplementedError

    def value(self, cx):
        raise NotImplementedError


# ------------------------------------------------------------------------------- builders
class SumIfs(Builder):
    """SUMIFS(data_row, year_row, year): every column whose year equals the period, e.g. the
    four quarters of a year on a quarterly grid. Several rows are added term by term."""
    fn = "sumifs"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g, self.rs = self.grid(), self.rows()
        self.scale, self.sign = self.scale_sign()
        self.year = p.get("year", "$year")

    def needs(self, cx):
        yield self.g.need(self.g.year_row)
        for r in self.rs:
            yield self.g.need(r)

    def formula(self, cx):
        y = cx.period(self.year).text()
        terms = [f"SUMIFS({self.g.row(r)},{self.g.years()},{y})" for r in self.rs]
        return _wrap("+".join(terms), self.scale, self.sign, len(terms) > 1)

    def value(self, cx):
        y, g = cx.period(self.year).year, self.g
        match = [c for c in g.cols()
                 if X.crit_equal_num(cx.model(g.sheet, g.year_row, c), y, f" ({g.sheet} row {g.year_row})")]
        total = 0.0
        for r in self.rs:
            s = 0.0
            for c in match:
                v = cx.model(g.sheet, r, c)
                if X.is_err(v):
                    return v
                if X.is_num(v):
                    s += v
            total += s
        return _finish(total, self.scale, self.sign)


class SumIfsLabel(Builder):
    """SUMIFS(INDEX(block,0,MATCH(year,year_row,0)), label_column, "label"): every row of a
    block whose label matches, in the column of the year. Row-move proof (rows found by
    label), for annual grids (MATCH takes the first column of the year)."""
    fn = "sumifs_label"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g = self.grid()
        self.label = p.get("label")
        if not isinstance(self.label, str) or not self.label or self.label[0] in "=<>":
            raise SpecError(f"{where}: 'label' must be non-empty text and not start with = < >")
        self.lcol = a1.col_index(p.get("label_col", "B"))
        self.r0, self.r1 = p.get("first_row"), p.get("last_row")
        if not (isinstance(self.r0, int) and isinstance(self.r1, int) and 0 < self.r0 <= self.r1):
            raise SpecError(f"{where}: first_row <= last_row required")
        self.scale, self.sign = self.scale_sign()
        self.year = p.get("year", "$year")

    def needs(self, cx):
        g = self.g
        yield g.need(g.year_row)
        yield (g.sheet, self.r0, self.r1, g.c0, g.c1)
        yield (g.sheet, self.r0, self.r1, self.lcol, self.lcol)

    def formula(self, cx):
        g, y = self.g, cx.period(self.year).text()
        crit = '"' + self.label.replace('"', '""') + '"'
        body = a1.rect_range(g.sheet, self.r0, self.r1, g.c0, g.c1)
        labels = a1.col_range(g.sheet, self.lcol, self.r0, self.r1)
        expr = f"SUMIFS(INDEX({body},0,MATCH({y},{g.years()},0)),{labels},{crit})"
        return _wrap(expr, self.scale, self.sign, False)

    def value(self, cx):
        g, y = self.g, cx.period(self.year).year
        pos = X.match_exact_num([cx.model(g.sheet, g.year_row, c) for c in g.cols()], y)
        if pos is None:
            return X.NA
        col = g.c0 + pos - 1
        total = 0.0
        for r in range(self.r0, self.r1 + 1):
            lab = cx.model(g.sheet, r, self.lcol)
            if not (isinstance(lab, str) and X.wildcard_match(self.label, lab)):
                continue
            v = cx.model(g.sheet, r, col)
            if X.is_err(v):
                return v
            if X.is_num(v):
                total += v
        return _finish(total, self.scale, self.sign)


class Window(Builder):
    """SUMPRODUCT((year_row>=y0)*(year_row<=y1)*data): a cumulative or windowed total.
    Omit y0 for "from the start of the grid"; y1 defaults to the column's year."""
    fn = "window"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g, self.rs = self.grid(), self.rows()
        self.scale, self.sign = self.scale_sign()
        self.y0 = p.get("y0")
        self.y1 = p.get("y1", "$year")

    def needs(self, cx):
        yield self.g.need(self.g.year_row)
        for r in self.rs:
            yield self.g.need(r)

    def formula(self, cx):
        g = self.g
        parts = []
        if self.y0 is not None:
            parts.append(f"({g.years()}>={cx.period(self.y0).text()})")
        if self.y1 is not None:
            parts.append(f"({g.years()}<={cx.period(self.y1).text()})")
        data = g.row(self.rs[0]) if len(self.rs) == 1 else "(" + "+".join(g.row(r) for r in self.rs) + ")"
        return _wrap(f"SUMPRODUCT({'*'.join(parts + [data])})", self.scale, self.sign, False)

    def value(self, cx):
        g = self.g
        y0 = cx.period(self.y0).year if self.y0 is not None else None
        y1 = cx.period(self.y1).year if self.y1 is not None else None
        total = 0.0
        err = None
        for c in g.cols():
            yv = cx.model(g.sheet, g.year_row, c)
            flags = []
            if y0 is not None:
                flags.append(X.cmp(yv, y0, ">="))
            if y1 is not None:
                flags.append(X.cmp(yv, y1, "<="))
            d = None
            for r in self.rs:
                x = X.arith(cx.model(g.sheet, r, c), f" ({g.sheet}!{a1.addr(c, r)})")
                d = x if d is None else X.add(d, x)
            e = next((f for f in flags if X.is_err(f)), None) or (d if X.is_err(d) else None)
            if e is not None:
                err = err or e
                continue
            f = 1.0
            for fl in flags:
                f *= 1.0 if fl else 0.0
            total += f * d
        if err is not None:
            return err
        return _finish(total, self.scale, self.sign)


class ColCount(Builder):
    """SUMPRODUCT((year_row>=y0)*(year_row<=y1))-expected: proves a stated window catches
    exactly the columns it claims (no gap, no duplicate, no stray text year)."""
    fn = "colcount"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g = self.grid()
        self.y0, self.y1, self.expect = p.get("y0"), p.get("y1"), p.get("expect")
        if self.y0 is None or self.y1 is None or not isinstance(self.expect, int):
            raise SpecError(f"{where}: colcount needs y0, y1 and expect (int)")

    def needs(self, cx):
        yield self.g.need(self.g.year_row)

    def formula(self, cx):
        yrs = self.g.years()
        return (f"SUMPRODUCT(({yrs}>={cx.period(self.y0).text()})*({yrs}<={cx.period(self.y1).text()}))"
                f"-{self.expect}")

    def value(self, cx):
        g = self.g
        y0, y1 = cx.period(self.y0).year, cx.period(self.y1).year
        n = 0.0
        for c in g.cols():
            yv = cx.model(g.sheet, g.year_row, c)
            a, b = X.cmp(yv, y0, ">="), X.cmp(yv, y1, "<=")
            for f in (a, b):
                if X.is_err(f):
                    return f
            n += 1.0 if (a and b) else 0.0
        return n - self.expect


class IndexMatch(Builder):
    """INDEX(data_row,1,MATCH(year,year_row,0)+offset): one column of the year. On a
    quarterly grid, offset 3 reads the fourth quarter (a stock at year end)."""
    fn = "index_match"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g = self.grid()
        self.r = self.rows()
        if len(self.r) != 1:
            raise SpecError(f"{where}: index_match takes one row")
        self.r = self.r[0]
        self.offset = p.get("offset", 0)
        if not isinstance(self.offset, int):
            raise SpecError(f"{where}: offset must be an int")
        self.scale, self.sign = self.scale_sign()
        self.year = p.get("year", "$year")

    def needs(self, cx):
        yield self.g.need(self.g.year_row)
        yield self.g.need(self.r)

    def formula(self, cx):
        m = f"MATCH({cx.period(self.year).text()},{self.g.years()},0)"
        if self.offset:
            m += f"{self.offset:+d}"
        return _wrap(f"INDEX({self.g.row(self.r)},1,{m})", self.scale, self.sign, False)

    def value(self, cx):
        g = self.g
        pos = X.match_exact_num([cx.model(g.sheet, g.year_row, c) for c in g.cols()],
                                cx.period(self.year).year)
        if pos is None:
            return X.NA
        k = pos + self.offset
        if not 1 <= k <= g.c1 - g.c0 + 1:
            return X.REF
        return _finish(cx.model(g.sheet, self.r, g.c0 + k - 1), self.scale, self.sign)


class YearEnd(Builder):
    """LOOKUP(2,1/(year_row=year),data_row): the LAST column of the year. Use it for a stock
    (a balance, a customer base, a network footprint); summing a stock over four quarters
    returns four times the stock and looks plausible."""
    fn = "year_end"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.g = self.grid()
        self.r = self.rows()
        if len(self.r) != 1:
            raise SpecError(f"{where}: year_end takes one row")
        self.r = self.r[0]
        self.scale, self.sign = self.scale_sign()
        self.year = p.get("year", "$year")

    def needs(self, cx):
        yield self.g.need(self.g.year_row)
        yield self.g.need(self.r)

    def formula(self, cx):
        y = cx.period(self.year).text()
        return _wrap(f"LOOKUP(2,1/({self.g.years()}={y}),{self.g.row(self.r)})",
                     self.scale, self.sign, False)

    def value(self, cx):
        g, y = self.g, cx.period(self.year).year
        last = None
        for c in g.cols():
            if X.cmp(cx.model(g.sheet, g.year_row, c), y, "=") is True:
                last = c
        if last is None:
            return X.NA
        return _finish(cx.model(g.sheet, self.r, last), self.scale, self.sign)


class Link(Builder):
    """A plain link, =Sheet!$C$5 (optionally scaled or negated). `refs` gives one address per
    year when the source is not on a grid. Links into the feed sheet use the sheet's
    CURRENT row numbers; the builder re-points them past any rows the plan inserts."""
    fn = "link"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.scale, self.sign = self.scale_sign()
        self.ref, self.refs = p.get("ref"), p.get("refs")
        if (self.ref is None) == (self.refs is None):
            raise SpecError(f"{where}: link needs exactly one of 'ref' or 'refs'")
        try:
            if self.ref is not None:
                self.target = a1.parse_cell_ref(self.ref)
            else:
                self.targets = {int(k): a1.parse_cell_ref(v) for k, v in self.refs.items()}
        except ValueError as e:
            raise SpecError(f"{where}: {e}") from None
        self.year = p.get("year", "$year")

    def _t(self, cx):
        if self.ref is not None:
            return self.target
        y = cx.period(self.year).year
        if y not in self.targets:
            raise SpecError(f"{self.where}: no ref for year {y}")
        return self.targets[y]

    def needs(self, cx):
        s, c, r = self._t(cx)
        yield (s, r, r, c, c)

    def formula(self, cx):
        s, c, r = self._t(cx)
        return _wrap(a1.cell_ref(s, c, cx.formula_row(s, r)), self.scale, self.sign, False)

    def value(self, cx):
        s, c, r = self._t(cx)
        return _finish(cx.model(s, r, c), self.scale, self.sign)


class SumRows(Builder):
    """SUM(C12,C13,C15): rows of this block (or "block.key") in the same column."""
    fn = "sum_rows"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.keys = p.get("keys")
        if not self.keys or not all(isinstance(k, str) for k in self.keys):
            raise SpecError(f"{where}: sum_rows needs 'keys'")

    def deps(self):
        return list(self.keys)

    def formula(self, cx):
        return "SUM(" + ",".join(cx.key_addr(k) for k in self.keys) + ")"

    def value(self, cx):
        return X.sum_refs([cx.key_value(k) for k in self.keys])


def _operand(spec, ctx, where):
    """A key (str) or an inline builder (dict with 'fn')."""
    if isinstance(spec, str):
        return ("key", spec)
    if isinstance(spec, dict) and "fn" in spec:
        return ("builder", make(spec, ctx, where))
    raise SpecError(f"{where}: operand must be a row key or a builder dict, got {spec!r}")


def _op_formula(op, cx):
    kind, x = op
    return cx.key_addr(x) if kind == "key" else "(" + x.formula(cx) + ")"


def _op_value(op, cx):
    kind, x = op
    return cx.key_value(x) if kind == "key" else x.value(cx)


def _op_deps(op):
    kind, x = op
    return [x] if kind == "key" else list(x.deps())


def _op_needs(op, cx):
    kind, x = op
    return () if kind == "key" else x.needs(cx)


class Combine(Builder):
    """A linear combination, left to right: C12-C15, C12-(SUMIFS(...)), 0.5*C12+C13.
    terms: [{"key": "total", "coef": 1}, {"fn": "sumifs", ..., "coef": -1}]"""
    fn = "combine"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        terms = p.get("terms")
        if not terms:
            raise SpecError(f"{where}: combine needs 'terms'")
        self.terms = []
        for i, t in enumerate(terms):
            if not isinstance(t, dict):
                raise SpecError(f"{where}: term {i} must be an object")
            coef = t.get("coef", 1)
            if not isinstance(coef, (int, float)) or isinstance(coef, bool) or coef == 0:
                raise SpecError(f"{where}: term {i} coef must be a non-zero number")
            spec = t["key"] if "key" in t else {k: v for k, v in t.items() if k != "coef"}
            self.terms.append((float(coef), _operand(spec, ctx, f"{where} term {i}")))

    def deps(self):
        return [d for _, op in self.terms for d in _op_deps(op)]

    def needs(self, cx):
        for _, op in self.terms:
            yield from _op_needs(op, cx)

    def formula(self, cx):
        out = []
        for i, (coef, op) in enumerate(self.terms):
            x = _op_formula(op, cx)
            mag = abs(coef)
            body = x if mag == 1 else f"{_num(mag)}*{x}"
            sign = "-" if coef < 0 else ("" if i == 0 else "+")
            out.append(sign + body)
        return "".join(out)

    def value(self, cx):
        acc = None
        for i, (coef, op) in enumerate(self.terms):
            v = _op_value(op, cx)
            if len(self.terms) == 1 and coef == 1:
                return 0.0 if v is None else v
            x = X.arith(v)
            mag = abs(coef)
            if mag != 1:
                x = X.mul(mag, x) if not X.is_err(x) else x
            if i == 0:
                acc = X.neg(x) if coef < 0 else x
            else:
                acc = X.sub(acc, x) if coef < 0 else X.add(acc, x)
        return acc


class Tie(Builder):
    """total - SUM(parts): the component rule's check row. total is a key or a builder."""
    fn = "tie"

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        if "total" not in p or not p.get("parts"):
            raise SpecError(f"{where}: tie needs 'total' and 'parts'")
        self.total = _operand(p["total"], ctx, f"{where} total")
        self.parts = list(p["parts"])

    def deps(self):
        return _op_deps(self.total) + self.parts

    def needs(self, cx):
        return _op_needs(self.total, cx)

    def formula(self, cx):
        return _op_formula(self.total, cx) + "-SUM(" + ",".join(cx.key_addr(k) for k in self.parts) + ")"

    def value(self, cx):
        return X.sub(X.arith(_op_value(self.total, cx)), X.sum_refs([cx.key_value(k) for k in self.parts]))


class Ratio(Builder):
    """num/den (times an optional multiplier). blank_if_zero writes IF(den=0,"",...): note
    that "" is NOT an empty cell to think-cell; prefer letting the error show."""
    fn = "ratio"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        if "num" not in p or "den" not in p:
            raise SpecError(f"{where}: ratio needs 'num' and 'den'")
        self.num = _operand(p["num"], ctx, f"{where} num")
        self.den = _operand(p["den"], ctx, f"{where} den")
        self.mult = p.get("multiplier", 1)
        self.blank = bool(p.get("blank_if_zero", False))

    def deps(self):
        return _op_deps(self.num) + _op_deps(self.den)

    def needs(self, cx):
        yield from _op_needs(self.num, cx)
        yield from _op_needs(self.den, cx)

    def formula(self, cx):
        n, d = _op_formula(self.num, cx), _op_formula(self.den, cx)
        q = f"{n}/{d}"
        if self.mult != 1:
            q = f"{q}*{_num(self.mult)}"
        return f'IF({d}=0,"",{q})' if self.blank else q

    def value(self, cx):
        dv = _op_value(self.den, cx)
        if self.blank:
            t = X.cmp(dv, 0, "=")
            if X.is_err(t):
                return t
            if t:
                return ""
        q = X.div(X.arith(_op_value(self.num, cx)), X.arith(dv))
        if self.mult != 1:
            q = X.mul(q, float(self.mult)) if not X.is_err(q) else q
        return q


class FyLabel(Builder):
    """Header label ="FY"&TEXT(MOD(year,100),"00") with an optional actual/forecast suffix
    taken from the model's own flag row: &IF(INDEX(flag,1,MATCH(year,years,0))=1,"A","F")."""
    fn = "fy_label"
    raw = True

    def __init__(self, p, ctx, where):
        super().__init__(p, ctx, where)
        self.year = p.get("year", "$year")
        self.flag = p.get("actual_flag")
        if self.flag:
            self.fg = self._grids.get(self.flag.get("grid"))
            if self.fg is None:
                raise SpecError(f"{where}: actual_flag.grid {self.flag.get('grid')!r} not defined")
            self.frow = self.flag.get("row")
            self.feq = self.flag.get("equals", 1)
            self.fsuf = self.flag.get("suffix", ["A", "F"])
            if not isinstance(self.frow, int) or len(self.fsuf) != 2:
                raise SpecError(f"{where}: actual_flag needs row (int) and suffix [actual, forecast]")

    def needs(self, cx):
        if self.flag:
            yield self.fg.need(self.fg.year_row)
            yield self.fg.need(self.frow)

    def formula(self, cx):
        y = cx.period(self.year).text()
        f = f'"FY"&TEXT(MOD({y},100),"00")'
        if self.flag:
            eq = f'"{self.feq}"' if isinstance(self.feq, str) else _num(self.feq)
            f += (f'&IF(INDEX({self.fg.row(self.frow)},1,MATCH({y},{self.fg.years()},0))={eq},'
                  f'"{self.fsuf[0]}","{self.fsuf[1]}")')
        return f

    def value(self, cx):
        y = cx.period(self.year).year
        s = "FY%02d" % (y % 100)
        if not self.flag:
            return s
        g = self.fg
        pos = X.match_exact_num([cx.model(g.sheet, g.year_row, c) for c in g.cols()], y)
        if pos is None:
            return X.NA
        v = cx.model(g.sheet, self.frow, g.c0 + pos - 1)
        t = X.cmp_eq_value(v, self.feq)
        if X.is_err(t):
            return t
        return s + (self.fsuf[0] if t else self.fsuf[1])


# ------------------------------------------------------------------ waterfall internals
class StepDefault(Builder):
    """Waterfall increment: (segment at close) - opening cell of the same row."""
    fn = "_step"

    def __init__(self, seg, open_col, where):
        self.seg, self.open_col, self.where = seg, open_col, where

    def needs(self, cx):
        return self.seg.needs(cx.with_period("close"))

    def formula(self, cx):
        return f"({self.seg.formula(cx.with_period('close'))})-{a1.addr(self.open_col, cx.row, True)}"

    def value(self, cx):
        return X.sub(X.arith(self.seg.value(cx.with_period("close"))),
                     X.arith(cx.cell_value(cx.row, self.open_col)))


class StepResidual(StepDefault):
    """Waterfall increment that closes the segment: (segment at close) - opening - other steps."""
    fn = "_residual"

    def __init__(self, seg, open_col, other_cols, where):
        super().__init__(seg, open_col, where)
        self.others = other_cols

    def formula(self, cx):
        f = super().formula(cx)
        if self.others:
            f += "-SUM(" + ",".join(a1.addr(c, cx.row) for c in self.others) + ")"
        return f

    def value(self, cx):
        v = super().value(cx)
        if self.others:
            v = X.sub(v, X.sum_refs([cx.cell_value(cx.row, c) for c in self.others]))
        return v


class RectCheck(Builder):
    """SUM(rectangle) - (model total at a period): opening stack / closing total checks.
    The rectangle is given as row OFFSETS inside the block (rows are placed later)."""
    fn = "_rect"

    def __init__(self, off0, off1, c0, c1, total, period_tok, where):
        self.off0, self.off1, self.c0, self.c1 = off0, off1, c0, c1
        self.total, self.tok, self.where = total, period_tok, where

    def _rows(self, cx):
        return cx.block.row_of(self.off0), cx.block.row_of(self.off1)

    def needs(self, cx):
        return self.total.needs(cx.with_period(self.tok))

    def formula(self, cx):
        r0, r1 = self._rows(cx)
        rng = f"{a1.addr(self.c0, r0)}:{a1.addr(self.c1, r1)}"
        return f"SUM({rng})-({self.total.formula(cx.with_period(self.tok))})"

    def value(self, cx):
        r0, r1 = self._rows(cx)
        cells = [cx.cell_value(r, c) for r in range(r0, r1 + 1) for c in range(self.c0, self.c1 + 1)]
        return X.sub(X.sum_refs(cells), X.arith(self.total.value(cx.with_period(self.tok))))


BUILDERS = {b.fn: b for b in (SumIfs, SumIfsLabel, Window, ColCount, IndexMatch, YearEnd, Link,
                              SumRows, Combine, Tie, Ratio, FyLabel)}


def make(spec, ctx, where):
    if not isinstance(spec, dict) or "fn" not in spec:
        raise SpecError(f"{where}: a formula must be an object with 'fn' (one of {sorted(BUILDERS)})")
    cls = BUILDERS.get(spec["fn"])
    if cls is None:
        raise SpecError(f"{where}: unknown fn {spec['fn']!r} (one of {sorted(BUILDERS)})")
    return cls(spec, ctx, where)
