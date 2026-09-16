"""Spec -> plan: every row and cell the build will write, at its final address.

Layout contract per standard block (offsets from the block's first row):
    title / notes (SOURCE -, BASIS -) / header (period labels) / SPACER / rows ... / gap
Waterfall block:
    title / notes / [period inputs] / header (open, steps, close) / SPACER / segments /
    blank / Check opening stack / Check closing total / gap

Anchors: {"row": n} (final row number), {"append": true} (below the last used row), or
{"insert_before": regex} (native Rows.Insert above the one row whose label matches; insert
blocks must come first in the spec). Row numbers anywhere else in the spec (donor rows,
links into the feed sheet) are the sheet's CURRENT numbers; the plan maps them.
"""
import re
from dataclasses import dataclass, field

from . import a1
from . import builders as B
from . import spec as S
from .builders import SpecError

ROW_KINDS = ("data", "memo", "check", "input")


@dataclass
class PCell:
    kind: str                       # formula | text | number
    builder: object = None
    periods: dict = None
    text: str = None
    number: float = None
    style: str = None               # "input" for typed inputs (blue)
    formula: str = None             # rendered, with the leading "="


@dataclass
class PRow:
    offset: int
    kind: str                       # title note inputs header spacer data memo check input blank gap
    label: str = None
    key: str = None
    fmt: str = None
    row: int = 0


@dataclass
class PBlock:
    id: str
    type: str
    spec: dict
    label_col: int
    first_col: int
    last_col: int
    rows: list = field(default_factory=list)
    first_row: int = 0
    anchor: str = ""
    insert_time_row: int = None     # where Rows.Insert happens (coordinates at that moment)
    insert_orig_row: int = None     # the anchor row in the sheet as it is now
    anchor_regex: str = None
    anchor_col: int = None
    donors: dict = field(default_factory=dict)   # kind -> CURRENT row number
    cells: dict = field(default_factory=dict)    # (offset, col) -> PCell

    @property
    def height(self):
        return len(self.rows)

    @property
    def last_row(self):
        return self.first_row + self.height - 1

    def row_of(self, offset):
        return self.first_row + offset

    def kind_rows(self, *kinds):
        return [self.row_of(r.offset) for r in self.rows if r.kind in kinds]

    def header_row(self):
        return self.kind_rows("header")[0]

    def chart_range(self):
        """The range to select in think-cell: header row down to the last series row."""
        series = self.kind_rows("data") if self.type == "standard" else self.kind_rows("segment")
        last = max(series) if series else self.header_row() + 1
        return f"{a1.addr(self.label_col, self.header_row())}:{a1.addr(self.last_col, last)}"


class SheetState:
    """The feed sheet as it is now. `values`: {(r, c): cached value}; `occupied`: every cell
    holding a value or a formula (a formula caching "" still occupies its cell)."""

    def __init__(self, values, occupied):
        self.cells = values
        self.occupied = set(occupied)
        self.last_row = max((r for r, _ in self.occupied), default=0)

    def find(self, regex, col):
        rx = re.compile(regex)
        return sorted(r for (r, c), v in self.cells.items()
                      if c == col and isinstance(v, str) and rx.search(v))

    def row_empty(self, r, c0, c1):
        return all((r, c) not in self.occupied for c in range(c0, c1 + 1))


class Ctx:
    """Formula/value context for one cell."""

    def __init__(self, plan, block, row, col, periods, model=None, evaluator=None):
        self.plan, self.block, self.row, self.col = plan, block, row, col
        self.periods = periods or {}
        self._model, self._eval = model, evaluator

    def period(self, tok):
        return B.resolve_period(tok, self.periods, f"{self.block.id}!{a1.addr(self.col, self.row)}")

    def with_period(self, name):
        p = dict(self.periods)
        p["year"] = self.periods[name]
        return Ctx(self.plan, self.block, self.row, self.col, p, self._model, self._eval)

    def key_row(self, key):
        bid, k = key.split(".", 1) if "." in key else (self.block.id, key)
        if (bid, k) not in self.plan.keys:
            raise SpecError(f"{self.block.id}: unknown row key {key!r}")
        return self.plan.keys[(bid, k)]

    def key_addr(self, key):
        return a1.addr(self.col, self.key_row(key))

    def key_value(self, key):
        return self.cell_value(self.key_row(key), self.col)

    def cell_value(self, r, c):
        if self._eval is None:
            raise RuntimeError("no evaluator: formula rendering only")
        return self._eval(r, c)

    def model(self, sheet, r, c):
        return self._model(sheet, r, c)

    def formula_row(self, sheet, r):
        return self.plan.to_new(r) if sheet.lower() == self.plan.feed_sheet.lower() else r


def _grids(spec):
    out = {}
    for name, g in (spec.get("grids") or {}).items():
        try:
            out[name] = B.Grid(name, g["sheet"], int(g["year_row"]), a1.col_index(g["first_col"]),
                               a1.col_index(g["last_col"]))
        except (KeyError, ValueError, TypeError) as e:
            raise SpecError(f"grid {name!r}: needs sheet, year_row, first_col, last_col ({e})") from None
        if out[name].c1 < out[name].c0:
            raise SpecError(f"grid {name!r}: last_col before first_col")
    return out


class Plan:
    def __init__(self, spec, state):
        self.spec = spec
        self.workbook = spec["workbook"]
        self.feed_sheet = spec["feed_sheet"]
        self.opts = spec["options"]
        self.grids = _grids(spec)
        self.ctx = {"grids": self.grids}
        self.state = state
        self.blocks = []
        self.keys = {}
        self.shifts = []            # (orig_row, n): rows >= orig_row move down by n
        self.cells = {}             # (row, col) -> (PBlock, PCell)
        self.errors, self.warnings = [], []
        self._place_all()
        if not self.errors:
            self._finalise()
        self.errors = list(dict.fromkeys(self.errors))
        self.warnings = list(dict.fromkeys(self.warnings))

    # ------------------------------------------------------------------ coordinates
    def to_new(self, r):
        return r + sum(n for at, n in self.shifts if r >= at)

    def to_orig(self, r, allow_new=False):
        """Final row -> current row, or None when the row belongs to a new block (with
        allow_new: None only for rows that a Rows.Insert creates)."""
        if not allow_new:
            for b in self.blocks:
                if b.first_row <= r <= b.last_row:
                    return None
        for cand in range(max(1, r - sum(n for _, n in self.shifts)), r + 1):
            if self.to_new(cand) == r:
                return cand
        return None

    # ------------------------------------------------------------------ placement
    def _place_all(self):
        seen_other = False
        ids = set()
        for i, bs in enumerate(self.spec["blocks"]):
            where = f"block {i} ({bs.get('id', '?')})"
            try:
                bid = bs.get("id")
                if not isinstance(bid, str) or not re.fullmatch(r"[A-Za-z0-9_\-]+", bid):
                    raise SpecError(f"{where}: 'id' must be letters, digits, _ or -")
                if bid in ids:
                    raise SpecError(f"{where}: duplicate id")
                ids.add(bid)
                anchor = bs.get("anchor") or {}
                is_insert = "insert_before" in anchor
                if is_insert and seen_other:
                    raise SpecError(f"{where}: insert_before blocks must come before append/row blocks")
                seen_other = seen_other or not is_insert
                blk = self._shape(bs, where)
                self._anchor(blk, anchor, where)
                self.blocks.append(blk)
            except SpecError as e:
                self.errors.append(str(e))

    def _shape(self, bs, where):
        typ = bs.get("type", "standard")
        lc = a1.col_index(bs.get("label_col", self.spec.get("label_col", "B")))
        fc = a1.col_index(bs.get("first_col", self.spec.get("first_col", "C")))
        if fc <= lc:
            raise SpecError(f"{where}: first_col must be right of label_col")
        title = bs.get("title")
        if not isinstance(title, str) or not title.strip():
            raise SpecError(f"{where}: 'title' required")
        notes = bs.get("notes") or []
        if not notes:
            self.warnings.append(f"{where}: no notes; state the SOURCE and the BASIS on the block")
        elif not any(n.upper().startswith("SOURCE") for n in notes) or not any(
                n.upper().startswith("BASIS") for n in notes):
            self.warnings.append(f"{where}: notes should include a 'SOURCE -' and a 'BASIS -' line")
        for t in [title] + list(notes):
            S.lint_text(t, where, self.errors)
        if typ == "standard":
            return self._shape_standard(bs, where, lc, fc, title, notes)
        if typ == "waterfall":
            return self._shape_waterfall(bs, where, lc, fc, title, notes)
        raise SpecError(f"{where}: type must be 'standard' or 'waterfall'")

    def _donors(self, bs, where):
        d = bs.get("donor") or {}
        ok = {"title", "note", "inputs", "header", "data", "memo", "check", "input", "blank",
              "segment", "spacer"}
        bad = set(d) - ok
        if bad:
            raise SpecError(f"{where}: unknown donor kinds {sorted(bad)}")
        for k, v in d.items():
            if not isinstance(v, int) or v < 1:
                raise SpecError(f"{where}: donor {k} must be a row number")
        return dict(d)

    def _columns(self, bs, where):
        cols = bs.get("columns")
        per = bs.get("periods")
        if (cols is None) == (per is None):
            raise SpecError(f"{where}: give exactly one of 'periods' or 'columns'")
        if per is not None:
            years = per.get("years")
            if years is None and "from" in per and "to" in per:
                years = list(range(int(per["from"]), int(per["to"]) + 1))
            if not years or not all(isinstance(y, int) for y in years):
                raise SpecError(f"{where}: periods needs 'years' or 'from'/'to'")
            tmpl = per.get("label", "FY{yy}")
            cols = [{"year": y, "label": tmpl} for y in years]
            header = per.get("header", "text")
            flag = per.get("actual_flag")
        else:
            header, flag = bs.get("header", "text"), bs.get("actual_flag")
        if header not in ("text", "formula"):
            raise SpecError(f"{where}: header must be 'text' or 'formula'")
        out = []
        for j, c in enumerate(cols):
            if not isinstance(c, dict) or not isinstance(c.get("year"), int):
                raise SpecError(f"{where}: column {j} needs an int 'year' (the window end for a window column)")
            frm = c.get("from")
            if frm is not None and (not isinstance(frm, int) or frm > c["year"]):
                raise SpecError(f"{where}: column {j} 'from' must be an int <= 'year'")
            label = c.get("label", "FY{yy}")
            try:
                text = label.format(yyyy=c["year"], yy="%02d" % (c["year"] % 100), y=c["year"],
                                    **({"from": frm, "ff": "%02d" % (frm % 100)} if frm else {}))
            except (KeyError, IndexError, ValueError) as e:
                raise SpecError(f"{where}: column {j} label {label!r}: {e!r}") from None
            out.append({"year": c["year"], "from": frm, "label": text,
                        "formula": header == "formula" and frm is None, "flag": flag})
            if not S.states_period(text):
                self.errors.append(f"{where}: column label {text!r} does not state its period")
            S.lint_text(text, where, self.errors)
        return out

    def _shape_standard(self, bs, where, lc, fc, title, notes):
        cols = self._columns(bs, where)
        blk = PBlock(bs["id"], "standard", bs, lc, fc, fc + len(cols) - 1)
        blk.donors = self._donors(bs, where)
        default_fmt = S.fmt(bs.get("format", "number"), where)
        add = blk.rows.append
        add(PRow(0, "title", title))
        for n in notes:
            add(PRow(len(blk.rows), "note", n))
        hdr = len(blk.rows)
        add(PRow(hdr, "header", bs.get("header_label")))
        for j, c in enumerate(cols):
            per = {"year": B.Period(c["year"])}
            if c["formula"]:
                fb = B.FyLabel({"fn": "fy_label", "actual_flag": c["flag"]}, self.ctx, f"{where} header")
                blk.cells[(hdr, fc + j)] = PCell("formula", fb, per)
            else:
                blk.cells[(hdr, fc + j)] = PCell("text", text=c["label"])
        add(PRow(len(blk.rows), "spacer"))
        seen_below = False
        for i, rs in enumerate(bs.get("rows") or []):
            rw = f"{where} row {i}"
            if rs.get("blank"):
                add(PRow(len(blk.rows), "blank"))
                continue
            kind = rs.get("kind", "data")
            if kind not in ROW_KINDS:
                raise SpecError(f"{rw}: kind must be one of {ROW_KINDS}")
            label = rs.get("label")
            if not isinstance(label, str) or not label.strip():
                raise SpecError(f"{rw}: 'label' required")
            S.lint_text(label, rw, self.errors)
            if kind == "check" and not (label.startswith("Check") and "(=0)" in label):
                raise SpecError(f"{rw}: a check row's label must start 'Check' and say '(=0)'")
            if kind in ("memo", "check"):
                seen_below = True
            elif kind == "data" and seen_below:
                self.warnings.append(f"{rw}: series row below memo/check rows; keep checks outside the chart range")
            key = rs.get("key")
            if key is not None and (not isinstance(key, str) or not re.fullmatch(r"[A-Za-z0-9_]+", key)):
                raise SpecError(f"{rw}: key must be letters, digits or _")
            fmtcode = S.fmt(rs.get("format", "check" if kind == "check" else default_fmt), rw)
            off = len(blk.rows)
            add(PRow(off, kind, label, key, fmtcode))
            only_first = bool(rs.get("only_first"))
            if kind == "input":
                vals = rs.get("values")
                if not rs.get("source"):
                    raise SpecError(f"{rw}: an input row needs 'source' (it is a typed value with no model link)")
                if not isinstance(vals, list) or len(vals) != len(cols):
                    raise SpecError(f"{rw}: input 'values' must list one value per column")
                for j, v in enumerate(vals):
                    if v is not None:
                        if not isinstance(v, (int, float)) or isinstance(v, bool):
                            raise SpecError(f"{rw}: input values must be numbers or null")
                        blk.cells[(off, fc + j)] = PCell("number", number=float(v), style="input")
                continue
            fspec = rs.get("formula")
            if fspec is None:
                raise SpecError(f"{rw}: 'formula' required (or kind 'input' with values)")
            bld = B.make(fspec, self.ctx, rw)
            for j, c in enumerate(cols):
                if only_first and j:
                    break
                per = {"year": B.Period(c["year"]),
                       "from": B.Period(c["from"]) if c["from"] else None}
                blk.cells[(off, fc + j)] = PCell("formula", bld, per)
        for _ in range(int(bs.get("gap_after", 1))):
            add(PRow(len(blk.rows), "gap"))
        if not blk.kind_rows("data"):
            self.warnings.append(f"{where}: no 'data' rows, nothing for a chart to select")
        return blk

    def _shape_waterfall(self, bs, where, lc, fc, title, notes):
        o, c = bs.get("open") or {}, bs.get("close") or {}
        if not isinstance(o.get("year"), int) or not isinstance(c.get("year"), int):
            raise SpecError(f"{where}: open.year and close.year (int) required")
        if o["year"] >= c["year"]:
            raise SpecError(f"{where}: open.year must be before close.year")
        segs = bs.get("segments") or []
        if len(segs) < 1:
            raise SpecError(f"{where}: 'segments' required")
        steps = bs.get("steps") or [{"label": s.get("label"), "segment": s.get("key")} for s in segs]
        total = bs.get("total")
        if not isinstance(total, dict):
            raise SpecError(f"{where}: 'total' (the model's own total, a builder) required for the check rows")
        nsteps = len(steps)
        close_col = fc + nsteps + 1
        blk = PBlock(bs["id"], "waterfall", bs, lc, fc, close_col)
        blk.donors = self._donors(bs, where)
        wfmt = S.fmt(bs.get("format", "waterfall"), where)
        if not wfmt.rstrip().endswith("@"):
            self.errors.append(f"{where}: a waterfall format must end ';;@' so the literal 'e' shows")
        add = blk.rows.append
        add(PRow(0, "title", title))
        for n in notes:
            add(PRow(len(blk.rows), "note", n))
        tmpl = bs.get("label", "FY{yy}")
        olab = o.get("label") or tmpl.format(yyyy=o["year"], yy="%02d" % (o["year"] % 100), y=o["year"])
        clab = c.get("label") or tmpl.format(yyyy=c["year"], yy="%02d" % (c["year"] % 100), y=c["year"])
        for lab in (olab, clab):
            if not S.states_period(lab):
                self.errors.append(f"{where}: open/close label {lab!r} does not state its period")
        per = {"open": B.Period(o["year"]), "close": B.Period(c["year"])}
        inputs = bool(bs.get("period_inputs"))
        if inputs:
            off = len(blk.rows)
            add(PRow(off, "inputs", bs.get("inputs_label", "Periods (inputs): opening year, closing year"),
                     fmt=S.FORMATS["year"]))
            blk.cells[(off, fc)] = PCell("number", number=float(o["year"]), style="input")
            blk.cells[(off, fc + 1)] = PCell("number", number=float(c["year"]), style="input")
            blk._inputs_offset = off
        hdr = len(blk.rows)
        add(PRow(hdr, "header", bs.get("header_label")))
        if inputs:
            for col, tok in ((fc, "open"), (close_col, "close")):
                fb = B.FyLabel({"fn": "fy_label", "year": "$" + tok}, self.ctx, f"{where} header")
                blk.cells[(hdr, col)] = PCell("formula", fb, dict(per))
        else:
            blk.cells[(hdr, fc)] = PCell("text", text=olab)
            blk.cells[(hdr, close_col)] = PCell("text", text=clab)
        for j, st in enumerate(steps):
            if not isinstance(st.get("label"), str) or not st["label"]:
                raise SpecError(f"{where}: step {j} needs a label")
            blk.cells[(hdr, fc + 1 + j)] = PCell("text", text=st["label"])
        add(PRow(len(blk.rows), "spacer"))
        seg_off = {}
        seg_bld = {}
        for i, sg in enumerate(segs):
            key, lab = sg.get("key"), sg.get("label")
            if not isinstance(key, str) or not re.fullmatch(r"[A-Za-z0-9_]+", key) or not isinstance(lab, str):
                raise SpecError(f"{where}: segment {i} needs key and label")
            off = len(blk.rows)
            add(PRow(off, "segment", lab, key, wfmt))
            seg_off[key] = off
            seg_bld[key] = B.make(sg.get("value"), self.ctx, f"{where} segment {key}")
            blk.cells[(off, fc)] = PCell("formula", seg_bld[key], dict(per, year=per["open"]))
        by_seg = {}
        for j, st in enumerate(steps):
            if st.get("segment") not in seg_off:
                raise SpecError(f"{where}: step {j} names unknown segment {st.get('segment')!r}")
            by_seg.setdefault(st["segment"], []).append(j)
        for key, js in by_seg.items():
            implicit = [j for j in js if not steps[j].get("formula")]
            if len(implicit) > 1:
                raise SpecError(f"{where}: segment {key} has {len(implicit)} steps without a formula; "
                                f"give all but one an explicit formula")
            for j in js:
                col = fc + 1 + j
                sw = f"{where} step {j}"
                if steps[j].get("formula"):
                    bld = B.make(steps[j]["formula"], self.ctx, sw)
                elif len(js) == 1:
                    bld = B.StepDefault(seg_bld[key], fc, sw)
                else:
                    bld = B.StepResidual(seg_bld[key], fc, [fc + 1 + k for k in js if k != j], sw)
                blk.cells[(seg_off[key], col)] = PCell("formula", bld, dict(per, year=per["close"]))
        last_seg = max(seg_off.values())
        blk.cells[(last_seg, close_col)] = PCell("text", text="e")
        add(PRow(len(blk.rows), "blank"))
        tot = B.make(total, self.ctx, f"{where} total")
        first_seg = min(seg_off.values())
        chk = S.FORMATS["check"]
        for tok, lab, c1 in (("open", f"Check (=0) - opening stack ties to the model total at {olab}", fc),
                             ("close", f"Check (=0) - opening plus steps ties to the model total at {clab}",
                              fc + nsteps)):
            off = len(blk.rows)
            add(PRow(off, "check", lab, None, chk))
            blk.cells[(off, fc)] = PCell("formula", B.RectCheck(first_seg, last_seg, fc, c1, tot, tok, where),
                                         dict(per, year=per[tok]))
        for _ in range(int(bs.get("gap_after", 1))):
            add(PRow(len(blk.rows), "gap"))
        blk._seg_off = seg_off
        return blk

    def _anchor(self, blk, anchor, where):
        if "insert_before" in anchor:
            col = a1.col_index(anchor.get("col", a1.col_letter(blk.label_col)))
            hits = self.state.find(anchor["insert_before"], col)
            if len(hits) != 1:
                raise SpecError(f"{where}: insert_before {anchor['insert_before']!r} matched rows {hits} "
                                f"in column {a1.col_letter(col)}; it must match exactly one")
            orig = hits[0]
            t = self.to_new(orig)
            for p in self.blocks:
                if p.first_row >= t:
                    p.first_row += blk.height
            blk.first_row = t
            blk.insert_time_row, blk.insert_orig_row = t, orig
            blk.anchor_regex, blk.anchor_col = anchor["insert_before"], col
            self.shifts.append((orig, blk.height))
            blk.anchor = f"insert {blk.height} rows above '{anchor['insert_before']}' (row {orig} now)"
        elif anchor.get("append"):
            last = max([self.to_new(self.state.last_row)] + [p.last_row for p in self.blocks])
            gap = int(anchor.get("gap", 1))
            blk.first_row = (last + 1 + gap) if last else 1
            blk.anchor = f"append below row {last}"
        elif isinstance(anchor.get("row"), int) and anchor["row"] > 0:
            blk.first_row = anchor["row"]
            blk.anchor = f"at row {anchor['row']}"
        else:
            raise SpecError(f"{where}: anchor must be {{'row': n}}, {{'append': true}} or {{'insert_before': regex}}")
        for p in self.blocks:
            if not (blk.last_row < p.first_row or blk.first_row > p.last_row):
                raise SpecError(f"{where}: rows {blk.first_row}-{blk.last_row} overlap block {p.id}")

    # ------------------------------------------------------------------ finalise
    def _finalise(self):
        for blk in self.blocks:
            for r in blk.rows:
                r.row = blk.row_of(r.offset)
                if r.key:
                    if (blk.id, r.key) in self.keys:
                        self.errors.append(f"{blk.id}: duplicate key {r.key!r}")
                    self.keys[(blk.id, r.key)] = r.row
            for k, v in blk.donors.items():
                if self.to_orig(self.to_new(v)) is None:
                    self.errors.append(f"{blk.id}: donor row {v} is not an existing row")
            if blk.insert_time_row is None:
                for r in range(blk.first_row, blk.last_row + 1):
                    o = self.to_orig(r, allow_new=True)
                    if o is not None and not self.state.row_empty(o, blk.label_col, blk.last_col):
                        self.errors.append(f"{blk.id}: target row {r} is not empty")
                        break
            for (off, col), cell in blk.cells.items():
                self.cells[(blk.row_of(off), col)] = (blk, cell)
        if self.errors:
            return
        for (r, c), (blk, cell) in sorted(self.cells.items()):
            if cell.kind != "formula":
                continue
            try:
                cell.formula = "=" + cell.builder.formula(self.cx(blk, r, c, cell))
            except SpecError as e:
                self.errors.append(str(e))
                continue
            if len(cell.formula) > 8000:
                self.errors.append(f"{blk.id}!{a1.addr(c, r)}: formula longer than Excel's 8,192 limit")

    def cx(self, blk, r, c, cell, model=None, evaluator=None):
        per = dict(cell.periods or {})
        if blk.type == "waterfall" and blk.spec.get("period_inputs"):
            ir = blk.row_of(blk._inputs_offset)
            for name, col in (("open", blk.first_col), ("close", blk.first_col + 1)):
                if per.get(name) is not None:
                    ref = a1.addr(col, ir, True)
                    per[name] = B.Period(per[name].year, ref)
            for name in ("open", "close"):
                if cell.periods and cell.periods.get("year") == cell.periods.get(name):
                    per["year"] = per[name]
        return Ctx(self, blk, r, c, per, model, evaluator)

    # ------------------------------------------------------------------ views
    def formula_cells(self):
        return [(r, c, blk, cell) for (r, c), (blk, cell) in sorted(self.cells.items()) if cell.kind == "formula"]

    def needs(self):
        out = set()
        for r, c, blk, cell in self.formula_cells():
            for n in cell.builder.needs(self.cx(blk, r, c, cell)):
                out.add(n)
        return sorted(out)

    def row_values(self, blk, prow):
        """{col: value-to-write} for one row, label column to last column."""
        vals = {c: None for c in range(blk.label_col, blk.last_col + 1)}
        if prow.label is not None and prow.kind not in ("spacer", "blank", "gap"):
            vals[blk.label_col] = prow.label
        for c in range(blk.first_col, blk.last_col + 1):
            cell = blk.cells.get((prow.offset, c))
            if cell is None:
                continue
            vals[c] = (cell.formula if cell.kind == "formula" else
                       cell.text if cell.kind == "text" else cell.number)
        return vals

    def empty_rows(self):
        """Every (row, c0, c1) that must stay blank: spacers, declared blanks, gaps."""
        return [(blk.row_of(r.offset), blk.label_col, blk.last_col, r.kind)
                for blk in self.blocks for r in blk.rows if r.kind in ("spacer", "blank", "gap")]

    def new_rows(self):
        s = set()
        for b in self.blocks:
            s.update(range(b.first_row, b.last_row + 1))
        return s

    def summary(self):
        lines = [f"PLAN  {self.workbook}", f"      sheet {self.feed_sheet!r}, last used row now {self.state.last_row}"]
        for b in self.blocks:
            lines.append(f"  block {b.id} [{b.type}] rows {b.first_row}-{b.last_row} "
                         f"cols {a1.col_letter(b.label_col)}:{a1.col_letter(b.last_col)}  ({b.anchor})")
            lines.append(f"    think-cell range {b.chart_range()}")
            for r in b.rows:
                row = b.row_of(r.offset)
                sample = ""
                cells = [(c, cl) for (o, c), cl in sorted(b.cells.items()) if o == r.offset]
                if cells:
                    c, cl = cells[0]
                    v = cl.formula if cl.kind == "formula" else cl.text if cl.kind == "text" else cl.number
                    if v is None:
                        v = "(formula not rendered: the plan has errors)"
                    sample = f"  {a1.addr(c, row)} {str(v)[:90]}" + (f"  (+{len(cells) - 1})" if len(cells) > 1 else "")
                lab = (r.label or "")[:48]
                lines.append(f"    r{row:<5} {r.kind:<8} {lab:<48}{sample}")
        for w in self.warnings:
            lines.append(f"  WARN  {w}")
        for e in self.errors:
            lines.append(f"  ERROR {e}")
        return "\n".join(lines)

    def to_json(self):
        blocks = []
        for b in self.blocks:
            blocks.append({
                "id": b.id, "type": b.type, "first_row": b.first_row, "last_row": b.last_row,
                "label_col": b.label_col, "first_col": b.first_col, "last_col": b.last_col,
                "chart_range": b.chart_range(), "anchor": b.anchor,
                "insert_time_row": b.insert_time_row, "insert_orig_row": b.insert_orig_row,
                "anchor_regex": b.anchor_regex, "anchor_col": b.anchor_col,
                "donors": {k: self.to_new(v) for k, v in b.donors.items()},
                "rows": [{"row": b.row_of(r.offset), "kind": r.kind, "label": r.label, "key": r.key,
                          "fmt": r.fmt} for r in b.rows],
                "cells": [{"row": b.row_of(o), "col": c, "kind": cl.kind, "formula": cl.formula,
                           "text": cl.text, "number": cl.number, "style": cl.style}
                          for (o, c), cl in sorted(b.cells.items())],
            })
        return {"workbook": self.workbook, "feed_sheet": self.feed_sheet, "shifts": self.shifts,
                "blocks": blocks, "warnings": self.warnings, "errors": self.errors}
