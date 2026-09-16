"""Lint rules added 2026-09-08 to close the gap with Witan's rule set (their ids in brackets).
Each rule: fn(wb, wbv_or_None, add) -> None; `add(level, rule, where, msg)`.
openpyxl only: `wb` holds formulas, `wbv` cached values (None in --quick mode)."""
import math
import re

from .analysis import tokens_of
from .describe import ftext, grid, iter_cells, peek
from .lint_rules import _cell_at, _col_width, _direct_refs_by_op, _fmt_class, _is_num, series_ref, split_ref_local

ARITH = {"+", "-", "*", "/", "^"}
AGG_FUNCS = {"SUM", "AVERAGE", "MIN", "MAX", "MEDIAN", "AVERAGEA", "SUMSQ"}
CUR_BRACKET = re.compile(r"\[\$([^\]\-]+)")
CUR_QUOTED = re.compile(r'"([^"]*[$€£¥][^"]*)"')
CUR_SYMBOL = re.compile(r"[$€£¥]")


def _val(cell):
    return None if cell is None else cell.value


def _calls(formula):
    """[(FUNCNAME, [arg_text, ...]), ...] for every function call in the formula (nested included)."""
    toks = tokens_of(formula)
    stack, out = [], []
    for t in toks:
        if t.type == "FUNC" and t.subtype == "OPEN":
            for fr in stack:
                fr[1][-1].append(t.value)
            stack.append([t.value[:-1].upper(), [[]]])
        elif t.type == "FUNC" and t.subtype == "CLOSE":
            if stack:
                name, args = stack.pop()
                out.append((name, ["".join(a).strip() for a in args]))
                for fr in stack:
                    fr[1][-1].append(")")
        elif t.type == "SEP" and t.subtype == "ARG" and stack:
            stack[-1][1].append([])
            for fr in stack[:-1]:
                fr[1][-1].append(",")
        else:
            for fr in stack:
                fr[1][-1].append(t.value)
    return out


def _range_cells(src, default_sheet, text):
    """Rows of cells (None where empty) of a plain range argument ('A2:B4', 'Sheet!$A$2:$B$4');
    (None, None) if not a plain range. Non-creating: see describe.grid."""
    text = text.strip()
    if not text or "(" in text or ":" not in text:
        return None, None
    try:
        sh, addr = split_ref_local(text, default_sheet)
        ws = src[sh] if sh in src.sheetnames else None
        if ws is None:
            for name in src.sheetnames:
                if name.strip().lower() == sh.strip().lower():
                    ws = src[name]
                    break
        if ws is None:
            return None, None
        return ws, grid(ws, addr)
    except Exception:
        return None, None


def _is_sorted(vals, descending=False):
    vals = [v for v in vals if v is not None]
    if len(vals) < 2:
        return True
    if all(_is_num(v) for v in vals):
        seq = vals
    elif all(isinstance(v, str) for v in vals):
        seq = [v.lower() for v in vals]
    else:
        return True  # mixed types: Excel's rules are subtle, do not guess
    pairs = zip(seq, seq[1:])
    return all((a >= b) if descending else (a <= b) for a, b in pairs)


def rule_lookups(wb, wbv, add, cap=40):
    """[D002] approximate-match lookup over an unsorted range; [D007] exact-match lookup over
    duplicate keys. VLOOKUP/HLOOKUP/MATCH/LOOKUP/XLOOKUP."""
    src = wbv or wb
    n_uns = n_dup = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and ("LOOKUP" in f.upper() or "MATCH(" in f.upper())):
                continue
            for name, args in _calls(f):
                approx = exact = descending = False
                key_axis = None  # ("col"|"row"|"vec", range_text)
                if name in ("VLOOKUP", "HLOOKUP") and len(args) >= 3:
                    mode = args[3].upper() if len(args) >= 4 else ""
                    approx = mode in ("", "TRUE", "1")
                    exact = mode in ("FALSE", "0")
                    key_axis = ("col" if name == "VLOOKUP" else "row", args[1])
                elif name == "MATCH" and len(args) >= 2:
                    mode = args[2].strip() if len(args) >= 3 else ""
                    approx = mode in ("", "1", "-1")
                    exact = mode == "0"
                    descending = mode == "-1"
                    key_axis = ("vec", args[1])
                elif name == "LOOKUP" and len(args) >= 2:
                    approx, key_axis = True, ("vec", args[1])
                elif name == "XLOOKUP" and len(args) >= 3:
                    mode = args[4].strip() if len(args) >= 5 else "0"
                    exact = mode in ("0", "")
                    approx = mode not in ("0", "")
                    key_axis = ("vec", args[1])
                if key_axis is None:
                    continue
                target_ws, cells = _range_cells(src, ws.title, key_axis[1])
                if cells is None:
                    continue
                if key_axis[0] == "row":
                    keys = [_val(x) for x in cells[0]]
                elif key_axis[0] == "col":
                    keys = [_val(row[0]) for row in cells]
                else:
                    keys = [_val(row[0]) for row in cells] if len(cells) >= len(cells[0]) else [_val(x) for x in cells[0]]
                keys = [k for k in keys if not ftext(k)]
                if approx and n_uns < cap and not _is_sorted(keys, descending=descending):
                    add("warn", "unsorted-lookup", f"{ws.title}!{c.coordinate}",
                        f"{name} approximate match over {key_axis[1]} which is not sorted: {f[:70]}")
                    n_uns += 1
                if exact and n_dup < cap:
                    seen, dups = set(), set()
                    for k in keys:
                        if k is None:
                            continue
                        kk = k.lower() if isinstance(k, str) else k
                        if kk in seen:
                            dups.add(k)
                        seen.add(kk)
                    if dups:
                        ex = ", ".join(repr(d) for d in list(dups)[:3])
                        add("warn", "duplicate-lookup-keys", f"{ws.title}!{c.coordinate}",
                            f"{name} exact match over {key_axis[1]} with duplicate keys ({ex}): returns the first")
                        n_dup += 1
            if n_uns >= cap and n_dup >= cap:
                return


def rule_aggregate_over_text(wb, wbv, add, cap=40):
    """[D005] SUM/AVERAGE/MIN/MAX over a range holding text or booleans, which are skipped silently."""
    if wbv is None:
        return
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and any(a + "(" in f.upper() for a in AGG_FUNCS)):
                continue
            for name, args in _calls(f):
                if name not in AGG_FUNCS:
                    continue
                for a in args:
                    tws, cells = _range_cells(wbv, ws.title, a)
                    if cells is None:
                        continue
                    bad = [(x.coordinate, x.value) for row in cells for x in row if x is not None
                           and ((isinstance(x.value, str) and x.value and not x.value.startswith("#"))
                                or isinstance(x.value, bool))]
                    if bad:
                        ex = ", ".join(f"{k}={v!r}" for k, v in bad[:3])
                        add("warn", "aggregate-over-text", f"{ws.title}!{c.coordinate}",
                            f"{name}({a}) skips {len(bad)} non-numeric cell(s): {ex}")
                        n += 1
                        break
                if n >= cap:
                    return


def _currency(fmt):
    if not fmt or fmt == "General":
        return None
    m = CUR_BRACKET.search(fmt)
    if m:
        return m.group(1).strip()
    m = CUR_QUOTED.search(fmt)
    if m:
        return m.group(1).strip()
    m = CUR_SYMBOL.search(fmt)
    return m.group(0) if m else None


def rule_currency_mixes(wb, wbv, add, cap=40):
    """[D008] one formula combining cells formatted in different currencies; [D023] + or -
    between a currency-formatted cell and a date/time-formatted cell."""
    n_cur = n_date = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not f:
                continue
            # D008: currencies among all directly referenced same-sheet cells and plain ranges
            symbols = {}
            for t in tokens_of(f):
                if t.type != "OPERAND" or t.subtype != "RANGE" or "!" in t.value:
                    continue
                addr = t.value.replace("$", "")
                try:
                    flat = [x for row in grid(ws, addr) for x in row if x is not None]
                except Exception:
                    continue
                for x in flat[:400]:
                    cur = _currency(x.number_format)
                    if cur and (_is_num(x.value) or ftext(x.value)):
                        symbols.setdefault(cur, x.coordinate)
            if len(symbols) >= 2 and n_cur < cap:
                ex = ", ".join(f"{a} ({s})" for s, a in list(symbols.items())[:3])
                add("error", "mixed-currency", f"{ws.title}!{c.coordinate}",
                    f"combines {len(symbols)} currencies: {ex}: {f[:60]}")
                n_cur += 1
            # D023: currency +/- date
            if n_date < cap and ("+" in f or "-" in f):
                for a1, a2 in _direct_refs_by_op(f):
                    try:
                        x1, x2 = _cell_at(ws, a1), _cell_at(ws, a2)
                    except Exception:
                        continue
                    if x1 is None or x2 is None:
                        continue
                    k1, k2 = _fmt_class(x1.number_format), _fmt_class(x2.number_format)
                    if {k1, k2} == {"currency", "date"}:
                        add("warn", "currency-date-mix", f"{ws.title}!{c.coordinate}",
                            f"{a1} ({k1}) +/- {a2} ({k2}): {f[:70]}")
                        n_date += 1
                        break
            if n_cur >= cap and n_date >= cap:
                return


def rule_broadcast_surprise(wb, wbv, add, cap=40):
    """[D006] a multi-cell range used as a direct operand of + - * / ^ outside any function:
    `=A1+B1:B10`. In a single cell this is implicit intersection (or a spill in 365)."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            f = ftext(c.value)
            if not (f and ":" in f and any(o in f for o in ARITH)):
                continue
            depth, seq = 0, []
            for t in tokens_of(f):
                if t.type == "FUNC":
                    depth += 1 if t.subtype == "OPEN" else -1
                    seq.append(("other", depth))
                elif t.type == "WHITE-SPACE":
                    continue
                elif depth == 0 and t.type == "OPERATOR-INFIX" and t.value in ARITH:
                    seq.append(("op", t.value))
                elif depth == 0 and t.type == "OPERAND" and t.subtype == "RANGE" and ":" in t.value:
                    seq.append(("range", t.value))
                else:
                    seq.append(("other", None))
            hit = None
            for i, (k, v) in enumerate(seq):
                if k == "op" and ((i > 0 and seq[i - 1][0] == "range") or (i + 1 < len(seq) and seq[i + 1][0] == "range")):
                    hit = seq[i - 1][1] if seq[i - 1][0] == "range" else seq[i + 1][1]
                    break
            if hit:
                add("info", "broadcast-surprise", f"{ws.title}!{c.coordinate}",
                    f"range {hit} used directly in arithmetic (implicit intersection / spill): {f[:70]}")
                n += 1
                if n >= cap:
                    return


def rule_row_height_clipping(wb, wbv, add, cap=30):
    """[D034] wrapped text that needs more lines than the row's explicit height allows."""
    n = 0
    for ws in wb.worksheets:
        for c in iter_cells(ws):
            v = c.value
            if not (isinstance(v, str) and v and not v.startswith("=")):
                continue
            al = c.alignment
            if not (al is not None and al.wrap_text):
                continue
            rd = ws.row_dimensions.get(c.row)
            height = rd.height if rd is not None else None
            if not height:
                continue  # auto height: Excel fits it
            width = _col_width(ws, c.column)
            lines = sum(max(1, math.ceil(len(part) / max(width, 1))) for part in v.split("\n"))
            need = lines * 15.0
            if need > height + 3:
                add("warn", "row-height-clipping", f"{ws.title}!{c.coordinate}",
                    f"~{lines} wrapped lines need ~{need:.0f} pt, row height {height:.0f} pt: text is cut")
                n += 1
                if n >= cap:
                    return


def _chart_px(ws, ch):
    """(width_px, height_px) from the anchor; falls back to the chart's cm size."""
    anc = getattr(ch, "anchor", None)
    frm = getattr(anc, "_from", None)
    to = getattr(anc, "to", None)
    ext = getattr(anc, "ext", None)
    if frm is not None and to is not None:
        w = sum(_col_width(ws, c + 1) * 7 for c in range(frm.col, max(to.col, frm.col)))
        w += (to.colOff - frm.colOff) / 9525 if hasattr(to, "colOff") else 0
        h = 0
        for r in range(frm.row, max(to.row, frm.row)):
            rd = ws.row_dimensions.get(r + 1)
            h += (rd.height if rd is not None and rd.height else 15) * 96 / 72
        h += (to.rowOff - frm.rowOff) / 9525 if hasattr(to, "rowOff") else 0
        return max(w, 0), max(h, 0)
    if ext is not None and getattr(ext, "cx", None) is not None:
        return ext.cx / 9525, ext.cy / 9525
    return (getattr(ch, "width", 15) or 0) * 37.8, (getattr(ch, "height", 7.5) or 0) * 37.8


def _ref_values(src, default_sheet, ref_f):
    ws, cells = _range_cells(src, default_sheet, ref_f)
    if cells is None:
        return []
    return [_val(x) for row in cells for x in row]


def rule_chart_geometry(wb, wbv, add, cap=20):
    """[D111-D116] chart-invisible (no visible area), chart-axis-labels-crowded (category labels
    cannot fit the plot width without rotation), chart-legend-crowded (too many series)."""
    src = wbv or wb
    n = 0
    for ws in wb.worksheets:
        for ci, ch in enumerate(getattr(ws, "_charts", []) or [], 1):
            where = f"{ws.title} chart{ci}"
            w, h = _chart_px(ws, ch)
            if w * h < 400:  # under ~20x20 px
                add("warn", "chart-invisible", where, f"anchored to ~{w:.0f}x{h:.0f} px: no visible area")
                n += 1
                continue
            series = getattr(ch, "series", []) or []
            if len(series) >= 8:
                add("info", "chart-legend-crowded", where,
                    f"{len(series)} series in one legend; group or split the chart")
                n += 1
            cats = None
            for ser in series:
                ref = series_ref(getattr(ser, "cat", None))
                if ref:
                    cats = _ref_values(src, ws.title, ref)
                if cats:
                    break
            if cats:
                labels = [str(v) for v in cats if v is not None]
                if labels:
                    longest = max(len(s) for s in labels)
                    rot = None
                    try:
                        rot = ch.x_axis.txPr.bodyPr.rot
                    except Exception:
                        rot = None
                    per_cat = (w * 0.85) / max(len(labels), 1)
                    if not rot and longest * 6.5 > per_cat and len(labels) >= 4:
                        add("info", "chart-axis-labels-crowded", where,
                            f"{len(labels)} category labels up to {longest} chars in ~{w:.0f} px "
                            f"(~{per_cat:.0f} px each): labels overlap or Excel rotates them")
                        n += 1
            if n >= cap:
                return


rule_aggregate_over_text.quick = False  # needs the values workbook

RULES = [rule_lookups, rule_aggregate_over_text, rule_currency_mixes, rule_broadcast_surprise,
         rule_row_height_clipping, rule_chart_geometry]
