"""Additional lint rules (openpyxl only). Each rule: fn(wb, wbv_or_None, add) -> None.
`add(level, rule, where, msg)`. Kept separate from lint.py so rules can grow independently.

Rule registry: every rule module exports RULES, and a rule that must not run in --quick mode
(needs the values workbook, or is slow) carries `quick = False`. lint.py builds the two lists
from that single declaration; before 2026-09-09 membership was declared in four places with a
circular import between the two rule modules.

Every read of a neighbouring cell goes through describe.peek / grid, never ws.cell() / ws[addr]:
those insert empty cells into the cached workbook permanently and inflate max_row/max_column.
Every formula test goes through describe.ftext so array formulas are inspected too.
"""
import re

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import range_boundaries

from .analysis import refs_in, tokens_of
from .describe import ftext, grid, grid_values, iter_cells, peek, peek_value, rows_of
from .wbcache import split_ref

SUM_RX = re.compile(r"\bSUM\s*\(", re.I)
# IF / IFS / IFERROR / IFNA as a function of their own; SUMIF, COUNTIFS, AVERAGEIF etc. do not count
IF_FUNC_RX = re.compile(r"(?<![A-Z_.])IF(S|ERROR|NA)?\s*\(", re.I)
LOCALE_TAG_RX = re.compile(r"\[\$-[0-9A-Fa-f]+\]")     # [$-409], [$-F800]: locale only, no symbol
QUOTED_RX = re.compile(r'"[^"]*"')
DATE_TOKEN_RX = re.compile(r"yy|dd|mmm|h:mm|\bd\b|\bm\b", re.I)


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _fmt_class(fmt):
    """percent | currency | date | number | general. Locale tags such as [$-409] carry no
    currency, and date tokens are tested before currency so `[$-F800]dddd, mmmm dd, yyyy`
    (Long Date) is a date, not a currency."""
    if not fmt or fmt == "General":
        return "general"
    f = LOCALE_TAG_RX.sub("", fmt).lower()
    if "%" in f:
        return "percent"
    has_currency = any(s in f for s in ("€", "$", "£", "eur", "usd", "gbp", "n$", "[$"))
    if has_currency:
        return "currency"
    if DATE_TOKEN_RX.search(QUOTED_RX.sub("", f)):  # literal text ("m", "days") is not a date token
        return "date"
    return "number"


def split_ref_local(ref, default_sheet):
    """'=Sheet!$A$1:$B$2', \"'Ana''s Sheet'!A1\" or 'A1:B2' -> (sheet, a1). Unescapes the doubled
    quote (the old version stripped only the outer quotes, so every range on a sheet with an
    apostrophe in its name was silently skipped by five rules)."""
    s, a = split_ref(ref.strip().lstrip("="))
    return (s if s is not None else default_sheet), a


def _col_width(ws, col):
    """Column width in characters, honouring min/max column-dimension spans; 8.43 default."""
    letter = get_column_letter(col)
    dim = ws.column_dimensions.get(letter)
    if dim is not None and dim.width:
        return dim.width
    for d in ws.column_dimensions.values():
        if d.min and d.max and d.min <= col <= d.max and d.width:
            return d.width
    return 8.43


def _direct_refs_by_op(formula):
    """[(addr1, addr2)] pairs of same-sheet single cells joined by + or - at any depth."""
    seq = []
    for t in tokens_of(formula):
        if t.type == "OPERATOR-INFIX" and t.value in ("+", "-"):
            seq.append(("op", t.value))
        elif t.type == "OPERAND" and t.subtype == "RANGE" and ":" not in t.value and "!" not in t.value:
            seq.append(("ref", t.value.replace("$", "")))
        elif t.type == "WHITE-SPACE":
            continue
        else:
            seq.append(("other", None))
    return [(seq[i - 1][1], seq[i + 1][1]) for i in range(1, len(seq) - 1)
            if seq[i][0] == "op" and seq[i - 1][0] == "ref" and seq[i + 1][0] == "ref"]


def series_ref(ref_obj):
    """The formula text behind a chart series part (val/cat/xVal/yVal): numRef first, strRef
    as fallback. One definition for both chart rules (they used to disagree on the order)."""
    if ref_obj is None:
        return None
    for kind in ("numRef", "strRef"):
        r = getattr(ref_obj, kind, None)
        if r is not None and getattr(r, "f", None):
            return r.f
    return None


def _cell_at(ws, addr):
    c, r, _, _ = range_boundaries(addr)
    return peek(ws, r, c)


def rule_sum_misses_adjacent(wb, wbv, add, cap=40):
    """SUM(B5:B9) where B10 (below) or B4 (above) holds a number contiguous with the range."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and SUM_RX.search(f)):
                continue
            for s, a, ext in refs_in(f, ws.title, wb):
                if ext or ":" not in a or s.lower() != ws.title.lower():
                    continue
                c1, r1, c2, r2 = range_boundaries(a)
                if None in (c1, r1, c2, r2):
                    continue
                if c1 == c2 and r2 > r1:  # vertical range
                    below = peek_value(ws, r2 + 1, c1)
                    if _is_num(below) and (r2 + 1, c1) != (c.row, c.column):
                        add("warn", "sum-misses-adjacent", f"{ws.title}!{c.coordinate}",
                            f"SUM({a}) stops at row {r2} but {get_column_letter(c1)}{r2 + 1} holds {below!r}")
                        n += 1
                elif r1 == r2 and c2 > c1:  # horizontal range
                    right = peek_value(ws, r1, c2 + 1)
                    if _is_num(right) and (r1, c2 + 1) != (c.row, c.column):
                        add("warn", "sum-misses-adjacent", f"{ws.title}!{c.coordinate}",
                            f"SUM({a}) stops at {get_column_letter(c2)} but {get_column_letter(c2 + 1)}{r1} holds {right!r}")
                        n += 1
                if n >= cap:
                    return


def rule_double_counting(wb, wbv, add, cap=40):
    """Two SUM ranges (or a SUM range and a direct ref) in one formula overlapping on one sheet."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and SUM_RX.search(f) and "+" in f):
                continue
            boxes = []
            for s, a, ext in refs_in(f, ws.title, wb):
                if ext:
                    continue
                try:
                    b = range_boundaries(a)
                except Exception:
                    continue
                if None in b:
                    continue
                boxes.append((s.lower(), b, a))
            for i in range(len(boxes)):
                for j in range(i + 1, len(boxes)):
                    (s1, (a1, b1, a2, b2), r1), (s2, (x1, y1, x2, y2), r2) = boxes[i], boxes[j]
                    if s1 == s2 and r1 != r2 and a1 <= x2 and a2 >= x1 and b1 <= y2 and b2 >= y1:
                        add("warn", "double-counting", f"{ws.title}!{c.coordinate}",
                            f"{r1} and {r2} overlap in the same formula: {f[:80]}")
                        n += 1
                        break
                else:
                    continue
                break
            if n >= cap:
                return


def rule_mixed_percent(wb, wbv, add, cap=40):
    """A + or - between a percent-formatted cell and a non-percent numeric cell."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and ("+" in f or "-" in f)):
                continue
            hit = None
            for a1, a2 in _direct_refs_by_op(f):
                try:
                    x1, x2 = _cell_at(ws, a1), _cell_at(ws, a2)
                except Exception:
                    continue
                if x1 is None or x2 is None:
                    continue
                if not all(_is_num(x.value) or ftext(x.value) for x in (x1, x2)):
                    continue
                k1, k2 = (_fmt_class(x.number_format) for x in (x1, x2))
                k1, k2 = ("number" if k == "general" else k for k in (k1, k2))
                if {k1, k2} == {"percent", "number"} or {k1, k2} == {"percent", "currency"}:
                    hit = (a1, k1, a2, k2)
                    break
            if hit:
                add("warn", "mixed-percent", f"{ws.title}!{c.coordinate}",
                    f"{hit[0]} ({hit[1]}) {'+/-'} {hit[2]} ({hit[3]}): {f[:70]}")
                n += 1
                if n >= cap:
                    return


def rule_empty_cell_coercion(wb, wbv, add, cap=40):
    """Arithmetic formula directly referencing an empty cell on the same sheet (reads as 0).
    Formulas guarded by IF/IFS/IFERROR/IFNA are exempt; SUMIF/COUNTIF are not a guard."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and any(op in f for op in "*/+-")):
                continue
            if IF_FUNC_RX.search(f):
                continue
            for s, a, ext in refs_in(f, ws.title, wb):
                if ext or ":" in a or s.lower() != ws.title.lower():
                    continue
                try:
                    x = _cell_at(ws, a)
                except Exception:
                    continue
                if x is None or x.value is None:
                    add("info", "empty-cell-coercion", f"{ws.title}!{c.coordinate}",
                        f"references empty {a} (coerces to 0): {f[:70]}")
                    n += 1
                    break
            if n >= cap:
                return


rule_empty_cell_coercion.quick = False


def rule_chart_series(wb, wbv, add, cap=40):
    """Chart series whose references point at a missing sheet, an empty range, or errors."""
    n = 0
    low = {ws.title.lower(): ws for ws in (wbv or wb).worksheets}
    for ws in wb.worksheets:
        for ci, ch in enumerate(getattr(ws, "_charts", []) or [], 1):
            for si, ser in enumerate(getattr(ch, "series", []) or [], 1):
                for part in ("val", "cat", "xVal", "yVal"):
                    ref = series_ref(getattr(ser, part, None))
                    if not ref:
                        continue
                    sheet, addr = split_ref_local(ref, ws.title)
                    target = low.get(sheet.lower())
                    where = f"{ws.title} chart{ci} series{si}.{part}"
                    if target is None:
                        add("error", "chart-series-broken", where, f"{ref}: sheet not found")
                        n += 1
                        continue
                    try:
                        flat = [v for row in grid_values(target, addr) for v in row]
                    except Exception:
                        add("error", "chart-series-broken", where, f"{ref}: bad range")
                        n += 1
                        continue
                    if all(v is None for v in flat):
                        add("warn", "chart-series-empty", where, f"{ref}: range is empty")
                        n += 1
                    elif any(isinstance(v, str) and v.startswith("#") for v in flat):
                        add("warn", "chart-series-errors", where, f"{ref}: contains error values")
                        n += 1
                    if n >= cap:
                        return


def rule_object_covers_cells(wb, wbv, add, cap=20):
    """Charts and images anchored over populated cells."""
    n = 0
    for ws in wb.worksheets:
        objs = [("chart", o) for o in getattr(ws, "_charts", []) or []] + [("image", o) for o in getattr(ws, "_images", []) or []]
        if not objs:
            continue
        occupied = {(c.row, c.column) for c in iter_cells(ws)}
        for kind, o in objs:
            anc = getattr(o, "anchor", None)
            frm = getattr(anc, "_from", None)
            to = getattr(anc, "to", None)
            if frm is None:
                continue
            r1, c1 = frm.row + 1, frm.col + 1
            if to is not None:
                r2, c2 = to.row + 1, to.col + 1
            else:
                r2, c2 = r1 + 15, c1 + 8
            covered = [(r, c) for (r, c) in occupied if r1 <= r <= r2 and c1 <= c <= c2]
            if covered:
                ex = ", ".join(f"{get_column_letter(c)}{r}" for r, c in sorted(covered)[:5])
                add("warn", "object-covers-cells", f"{ws.title} {kind}",
                    f"covers {len(covered)} populated cells ({get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}), e.g. {ex}")
                n += 1
                if n >= cap:
                    return


def _display_len(v, fmt):
    if isinstance(v, bool):
        return 5
    if isinstance(v, (int, float)):
        f = (fmt or "General")
        decimals = 0
        m = re.search(r"0\.(0+)", f)
        if m:
            decimals = len(m.group(1))
        scale = 1.0
        if f.endswith(",") or ",\"" in f or re.search(r"0,\s*$|0,\)", f):
            scale = 1e-3
        if f.count(",,"):
            scale = 1e-6
        if "%" in f:
            v = v * 100
        num = abs(v * scale)
        ip = len(f"{int(num):,}") if "#,##" in f or "0,0" in f else len(str(int(num)))
        return ip + (decimals + 1 if decimals else 0) + (1 if v < 0 else 0) + (1 if "%" in f else 0) + (2 if "(" in f else 0)
    return len(str(v))


def rule_column_clipping(wb, wbv, add, cap=40):
    """Numbers that will not fit their column width (Excel shows #####). Heuristic: one char
    per width unit; wrapped/text cells are ignored, numbers only."""
    n = 0
    src = wbv or wb
    for ws in src.worksheets:
        for c in iter_cells(ws):
            v = c.value
            if not _is_num(v):
                continue
            w = _col_width(ws, c.column)
            need = _display_len(v, c.number_format)
            if need > w + 0.5:
                add("warn", "column-clipping", f"{ws.title}!{c.coordinate}",
                    f"needs ~{need} chars, column width {w:.1f} (shows ##### ); format {c.number_format!r}")
                n += 1
                if n >= cap:
                    return


rule_column_clipping.quick = False


def rule_data_validation(wb, wbv, add, cap=40):
    """Cell values that violate their own data-validation list or numeric bounds."""
    n = 0
    src = wbv or wb
    for ws in wb.worksheets:
        dvs = getattr(ws, "data_validations", None)
        if not dvs or not dvs.dataValidation:
            continue
        wsv = src[ws.title]
        for dv in dvs.dataValidation:
            allowed = None
            if dv.type == "list" and dv.formula1:
                f1 = dv.formula1.strip()
                if f1.startswith('"') and f1.endswith('"'):
                    allowed = {x.strip() for x in f1.strip('"').split(",")}
                else:
                    try:
                        s, a = split_ref_local(f1, ws.title)
                        # grid() handles a single cell ($Z$1) and a range alike; the old
                        # `cells[0]` on a single Cell raised TypeError into a bare except
                        allowed = {str(v) for row in grid_values(src[s], a) for v in row if v is not None}
                    except Exception:
                        allowed = None
            lo = hi = None
            if dv.type in ("whole", "decimal") and dv.operator in (None, "between"):
                try:
                    lo, hi = float(dv.formula1), float(dv.formula2)
                except Exception:
                    lo = hi = None
            if allowed is None and lo is None:
                continue
            for rng in dv.sqref.ranges:
                # clamp to the populated extent: a validation on A:A spans 1,048,576 rows
                r2 = min(rng.max_row, wsv.max_row or 1)
                c2 = min(rng.max_col, wsv.max_column or 1)
                if r2 < rng.min_row or c2 < rng.min_col:
                    continue
                addr = f"{get_column_letter(rng.min_col)}{rng.min_row}:{get_column_letter(c2)}{r2}"
                for row in grid(wsv, addr):
                    for cell in row:
                        if cell is None or cell.value is None:
                            continue
                        v = cell.value
                        bad = False
                        if allowed is not None and str(v) not in allowed:
                            bad = True
                        if lo is not None and _is_num(v) and not (lo <= v <= hi):
                            bad = True
                        if bad:
                            add("error", "data-validation-breach", f"{ws.title}!{cell.coordinate}",
                                f"{v!r} violates {dv.type} rule {dv.formula1}{(' .. ' + str(dv.formula2)) if dv.formula2 else ''}")
                            n += 1
                            if n >= cap:
                                return


rule_data_validation.quick = False


def rule_row_format_inconsistent(wb, wbv, add, cap=30):
    """Numeric cells in one row with more than one number-format class (0.0% next to 0.00)."""
    n = 0
    for ws in wb.worksheets:
        for r, cells in rows_of(ws):
            nums = [c for c in cells if _is_num(c.value) or ftext(c.value)]
            if len(nums) < 4:
                continue
            classes = {}
            for c in nums:
                classes.setdefault(_fmt_class(c.number_format), []).append(c.coordinate)
            classes.pop("general", None)
            if len(classes) >= 2:
                minority = min(classes.items(), key=lambda kv: len(kv[1]))
                majority = max(classes.items(), key=lambda kv: len(kv[1]))
                maj_cols = [column_index_from_string(re.match(r"[A-Z]+", a).group()) for a in majority[1]]
                min_cols = [column_index_from_string(re.match(r"[A-Z]+", a).group()) for a in minority[1]]
                interior = all(min(maj_cols) < c < max(maj_cols) for c in min_cols)
                if interior and len(minority[1]) <= max(2, len(nums) // 5):
                    add("info", "row-format-inconsistent", f"{ws.title}!{minority[1][0]}",
                        f"{minority[0]} among {len(majority[1])} {majority[0]} cells in row {r}")
                    n += 1
                    if n >= cap:
                        return


def rule_external_link_roots(wb, wbv, add):
    from .paths import SYNC_ROOTS
    links = getattr(wb, "_external_links", None) or []
    roots = [str(r).lower() for r in SYNC_ROOTS]  # noqa: F841 - kept for a future same-root check
    for i, ln in enumerate(links, 1):
        tgt = str(getattr(getattr(ln, "file_link", None), "Target", "") or "")
        t = tgt.lower().replace("/", "\\")
        if t.startswith("file:") or ":\\" in t or t.startswith("\\\\") or "sharepoint" in t or "http" in t:
            add("warn", "external-link-absolute", f"[{i}]",
                f"{tgt[:90]} is absolute; breaks when the files move; put source and dependent in the same folder")


RULES = [rule_sum_misses_adjacent, rule_double_counting, rule_mixed_percent, rule_chart_series,
         rule_object_covers_cells, rule_row_format_inconsistent, rule_external_link_roots,
         rule_empty_cell_coercion, rule_column_clipping, rule_data_validation]
