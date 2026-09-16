"""Verification after a build, on the saved file. Everything but the last two checks is read
from the package (no Excel). Order follows what bites hardest.

  layout            spacer / blank / gap rows empty; labels on their rows; formula cells hold
                    formulas; period labels present and stating their period
  injected          every formula cell caches exactly the computed value
  checks_zero       every Check (=0) cell reads 0 (within options.check_tol)
  names             think-cell's sheet-scoped names: same set as the backup, re-pointed exactly
                    as the insertions require, none broken
  untouched_feed    every cell of the feed sheet outside the new blocks: same value as the
                    backup (moved rows compared at their new row); formula text too, with the
                    expected re-pointing applied
  untouched_model   every cached value on every other sheet unchanged (skipped above
                    options.deep_limit_mb unless --deep)
  errors            no new error values and no new '#REF!' in formula text on the feed sheet
  merged            no merged cells in the new rows
  calcpr            calcPr equals the backup's (plus the configured override)
  package           no [trash] parts left
  check_open        `xl check-open` exit 0 (the only check that catches an invalid package)
  engine (optional) on a throwaway copy, `xl calc --verify`: a full recalculation reproduces
                    every injected value (within tolerance); no error cells on the feed sheet
"""
import os
import re
import shutil
import time

from . import a1
from . import package as P
from . import reader
from . import xlbridge
from . import xlsem as X
from . import spec as S


class Result:
    def __init__(self):
        self.items = []

    def add(self, name, status, detail=""):
        self.items.append({"check": name, "status": status, "detail": detail})

    @property
    def failed(self):
        return [i for i in self.items if i["status"] in ("FAIL", "INCONCLUSIVE")]

    def text(self):
        return "\n".join(f"  {i['status']:<12} {i['check']:<16} {i['detail']}" for i in self.items)


def _fmt(items, n=6):
    s = "; ".join(str(x) for x in items[:n])
    return s + (f" ... (+{len(items) - n})" if len(items) > n else "")


def run(plan, backup, target, values, errors, expected_calcpr, *, session, deep=False,
        engine=False, open_check=True, log=print):
    res = Result()
    opts = plan.opts
    tol = float(opts["check_tol"])
    sheet = plan.feed_sheet
    new_rows = plan.new_rows()
    bvals, bforms = reader.scan_sheet(backup, sheet)
    tvals, tforms = reader.scan_sheet(target, sheet)

    # ---- layout
    probs = []
    for r, c0, c1, kind in plan.empty_rows():
        hit = [a1.addr(c, r) for c in range(c0, c1 + 1) if (r, c) in tvals or (r, c) in tforms]
        if hit:
            probs.append(f"{kind} row {r} not empty: {','.join(hit[:4])}")
    for b in plan.blocks:
        for pr in b.rows:
            r = b.row_of(pr.offset)
            if pr.label is not None and pr.kind not in ("spacer", "blank", "gap"):
                got = tvals.get((r, b.label_col))
                if got != pr.label:
                    probs.append(f"{a1.addr(b.label_col, r)} label {str(got)[:30]!r} != plan {pr.label[:30]!r}")
            for c in range(b.first_col, b.last_col + 1):
                cell = b.cells.get((pr.offset, c))
                if cell is None:
                    continue
                if cell.kind == "formula" and (r, c) not in tforms:
                    probs.append(f"{a1.addr(c, r)} should hold a formula, holds {tvals.get((r, c))!r}")
                if cell.kind == "text" and tvals.get((r, c)) != cell.text:
                    probs.append(f"{a1.addr(c, r)} text {tvals.get((r, c))!r} != {cell.text!r}")
                if cell.kind == "number" and not X.same_value(tvals.get((r, c)), cell.number):
                    probs.append(f"{a1.addr(c, r)} number {tvals.get((r, c))!r} != {cell.number!r}")
        hr = b.header_row()
        labels = [tvals.get((hr, c)) for c in range(b.first_col, b.last_col + 1)]
        period_cols = labels if b.type == "standard" else [labels[0], labels[-1]]
        bad = [l for l in period_cols if not (isinstance(l, str) and S.states_period(l))]
        if bad:
            probs.append(f"{b.id} header row {hr}: labels without a period {bad[:4]}")
        if any(isinstance(l, str) and "%%" in l for l in labels):
            probs.append(f"{b.id} header row {hr}: '%%' in a label")
    res.add("layout", "FAIL" if probs else "PASS",
            _fmt(probs) if probs else f"{len(plan.blocks)} blocks, {len(plan.empty_rows())} blank rows empty, "
                                      f"labels and period headers in place")

    # ---- injected values
    miss, diff = [], []
    for (r, c), want in values.items():
        if (r, c) in errors:
            continue
        got = tvals.get((r, c))
        if (r, c) not in tforms:
            miss.append(a1.addr(c, r))
        elif not X.same_value(got, want, rel=1e-12, abs_tol=1e-12):
            diff.append(f"{a1.addr(c, r)} cached {got!r} computed {want!r}")
    if errors:
        diff.extend(f"{a1.addr(c, r)} computes to {v} (not injected)" for (r, c), v in sorted(errors.items()))
    ok = not miss and not diff
    res.add("injected", "PASS" if ok else "FAIL",
            f"{len(values) - len(errors)} formula cells carry the computed value" if ok
            else _fmt(([f"no formula: {m}" for m in miss]) + diff))

    # ---- check rows
    checks, nonzero = 0, []
    for b in plan.blocks:
        for r in b.kind_rows("check"):
            for c in range(b.first_col, b.last_col + 1):
                if (r, c) not in tforms:
                    continue
                checks += 1
                v = tvals.get((r, c))
                if not (X.is_num(v) and abs(v) <= tol):
                    nonzero.append(f"{a1.addr(c, r)}={v!r}")
                if "IFERROR" in tforms[(r, c)].upper():
                    nonzero.append(f"{a1.addr(c, r)} wraps a check in IFERROR (a check must be able to fail)")
    res.add("checks_zero", "FAIL" if nonzero or not checks else "PASS",
            _fmt(nonzero) if nonzero else (f"{checks} check cells read 0 (tol {tol:g})" if checks
                                           else "no check cells found"))

    # ---- think-cell names
    nb, nt = P.local_names(backup, sheet), P.local_names(target, sheet)
    probs = []
    if set(nb) != set(nt):
        probs.append(f"names added {sorted(set(nt) - set(nb))}, lost {sorted(set(nb) - set(nt))}")
    for k in sorted(set(nb) & set(nt)):
        want = a1.shift_sheet_refs(nb[k]["text"], sheet, plan.to_new)
        if a1.norm_ref_text(want) != a1.norm_ref_text(nt[k]["text"]):
            probs.append(f"{k}: {nt[k]['text']} (expected {want})")
        if "#REF!" in nt[k]["text"]:
            probs.append(f"{k} is broken: {nt[k]['text']}")
        if nb[k]["hidden"] != nt[k]["hidden"]:
            probs.append(f"{k}: hidden flag changed")
    tc = [k for k in nb if k.lower().startswith("___thinkcell")]
    moved = [k for k in nb if a1.norm_ref_text(nb[k]["text"]) != a1.norm_ref_text(nt.get(k, {}).get("text", ""))]
    res.add("names", "FAIL" if probs else "PASS",
            _fmt(probs) if probs else f"{len(nb)} sheet-scoped names ({len(tc)} think-cell), "
                                      f"{len(moved)} re-pointed by the insertions as expected")

    # ---- untouched feed sheet
    diffs, fdiffs, compared = [], [], 0
    for (r, c), v in bvals.items():
        nr = plan.to_new(r)
        compared += 1
        if not X.same_value(v, tvals.get((nr, c))):
            diffs.append(f"{a1.addr(c, r)}->{a1.addr(c, nr)} {v!r} -> {tvals.get((nr, c))!r}")
    for (r, c), f in bforms.items():
        nr = plan.to_new(r)
        want = _shift_formula(f, sheet, plan)
        got = tforms.get((nr, c))
        if got is None or _norm_f(got) != _norm_f(want):
            fdiffs.append(f"{a1.addr(c, nr)} {got!r} expected {want!r}")
    spans = {}
    for b in plan.blocks:
        for r in range(b.first_row, b.last_row + 1):
            spans[r] = (b.label_col, b.last_col)
    extra = []
    for (r, c) in sorted(set(tvals) | set(tforms)):
        if r in spans:
            if not spans[r][0] <= c <= spans[r][1]:
                extra.append(a1.addr(c, r))
            continue
        o = plan.to_orig(r)
        if o is None or ((o, c) not in bvals and (o, c) not in bforms):
            extra.append(a1.addr(c, r))
    status = "FAIL" if diffs or fdiffs or extra else "PASS"
    detail = (_fmt(diffs + fdiffs + [f"unexpected cell {x}" for x in extra]) if status == "FAIL" else
              f"{compared} existing cells unchanged in value, {len(bforms)} existing formulas as expected"
              + (f" (rows below the insertion moved by {sum(n for _, n in plan.shifts)}, references "
                 f"re-pointed)" if plan.shifts else ""))
    res.add("untouched_feed", status, detail)

    # ---- untouched model sheets
    size_mb = os.path.getsize(target) / 1e6
    if deep or size_mb <= float(opts["deep_limit_mb"]):
        n, dd, nd = reader.compare_workbooks(backup, target, skip_sheets=[sheet])
        res.add("untouched_model", "FAIL" if nd else "PASS",
                _fmt([f"{s}!{a1.addr(c, r)} {a!r} -> {b!r}" for s, r, c, a, b in dd if s != "<sheets>"]
                     + [f"sheet list {d[1]} -> {d[2]}" for d in dd if d[0] == "<sheets>"])
                if nd else f"{n} cached values on the other sheets unchanged")
    else:
        res.add("untouched_model", "SKIPPED", f"{size_mb:.0f} MB > deep_limit_mb; run verify --deep to compare")

    # ---- errors and #REF!
    def errs(vals):
        return {k for k, v in vals.items() if X.is_err(v)}
    new_err = sorted(errs(tvals) - {(plan.to_new(r), c) for r, c in errs(bvals)})
    ref_b = sum(1 for f in bforms.values() if "#REF!" in f.upper())
    ref_t = sum(1 for f in tforms.values() if "#REF!" in f.upper())
    probs = [f"new error value at {a1.addr(c, r)}: {tvals[(r, c)]}" for r, c in new_err]
    if ref_t > ref_b:
        probs.append(f"formulas containing '#REF!': {ref_b} -> {ref_t}")
    res.add("errors", "FAIL" if probs else "PASS",
            _fmt(probs) if probs else f"no new error values; '#REF!' in formula text: {ref_t} (backup {ref_b})")

    # ---- merged cells
    merged = []
    for rg in P.merged_ranges(target, sheet):
        m = re.match(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", rg)
        r0, r1 = int(m.group(2)), int(m.group(4) or m.group(2))
        if any(r in new_rows for r in range(r0, r1 + 1)):
            merged.append(rg)
    res.add("merged", "FAIL" if merged else "PASS", _fmt(merged) if merged else "no merged cells in the new rows")

    # ---- calcPr
    got = xlbridge.cached().read_calcpr(target)
    res.add("calcpr", "PASS" if got == expected_calcpr else "FAIL",
            f"{got}" if got == expected_calcpr else f"{got} != expected {expected_calcpr}")
    warn = xlbridge.cached()._calcpr_warning(got)
    if warn:
        res.add("calcpr_warning", "WARN", warn)

    # ---- package
    trash = P.trash_parts(target)
    res.add("package", "FAIL" if trash else "PASS", _fmt(trash) if trash else "no [trash] parts")

    # ---- check-open
    if open_check:
        code, out = xlbridge.check_open(target, session)
        tail = " | ".join(ln.strip() for ln in out.splitlines()[-3:])
        res.add("check_open", "PASS" if code == 0 else "FAIL", f"exit {code}: {tail}"[:300])
    else:
        res.add("check_open", "SKIPPED", "not requested")

    # ---- engine check on a throwaway copy
    if engine:
        _engine(res, plan, target, values, errors, session, tol, log)
    return res


def _engine(res, plan, target, values, errors, session, tol, log):
    work = os.path.join(xlbridge.paths().WORK, "tcfeed_engine")
    os.makedirs(work, exist_ok=True)
    copy = os.path.join(work, f"{int(time.time() * 1000)}_{os.path.basename(target)}")
    shutil.copyfile(target, copy)
    try:
        code, rep, err = xlbridge.calc_verify(copy, session)
    finally:
        try:
            os.remove(copy)
        except OSError:
            pass
    if rep is None:
        res.add("engine", "FAIL", f"xl calc gave no JSON (exit {code}): {err[:200]}")
        return
    sheet = plan.feed_sheet.lower()
    new_rows = plan.new_rows()
    feed_err = [e for e in rep.get("errors", []) if e["sheet"].lower() == sheet
                and a1.split_cell(e["cell"])[1] in new_rows]
    real, fuzz, other = [], 0, 0
    for ch in rep.get("changed", []):
        if ch["sheet"].lower() != sheet:
            other += 1
            continue
        c, r = a1.split_cell(ch["cell"])
        want = values.get((r, c))
        after = ch["after"]
        if isinstance(after, int) and not isinstance(after, bool) and after < -2e9:
            after = X.XlError(str(after))
        if X.same_value(after, want, rel=1e-9, abs_tol=tol):
            fuzz += 1
        else:
            real.append(f"{ch['cell']} injected {ch['before']!r} recalculated {ch['after']!r}")
    capped = rep.get("changed_count", 0) >= 200
    if feed_err or real:
        status = "FAIL"
    elif capped:
        status = "INCONCLUSIVE"
    else:
        status = "PASS"
    detail = (_fmt([f"error {e['cell']} {e['error']}" for e in feed_err] + real) if status == "FAIL" else
              f"full rebuild reproduces all {len(values) - len(errors)} injected values"
              f" ({fuzz} within float tolerance); {other} changed cells elsewhere in the model")
    if capped:
        detail += "; changed list capped at 200 cells, result not conclusive"
    res.add("engine", status, detail)


_REF_ANY = re.compile(
    r"(?<![\w.\]'!$])"
    r"(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?"
    r"(?P<c1>\$?[A-Z]{1,3})(?P<r1>\$?\d+)"
    r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<r2>\$?\d+))?"
    r"(?![\w(!'])")


def _shift_formula(f, sheet, plan):
    """Expected text of an existing formula after the insertions: A1 references to the feed
    sheet (qualified with its name, or unqualified) have their rows mapped by plan.to_new;
    references to other sheets and external books are left alone. Whole-row/column
    references are not rewritten."""
    if not plan.shifts:
        return f

    def row(tok):
        return ("$" if tok.startswith("$") else "") + str(plan.to_new(int(tok.lstrip("$"))))

    def repl(m):
        s = m.group("sheet")
        if s is not None and a1.unquote_sheet(s).lower() != sheet.lower():
            return m.group(0)
        out = (s + "!" if s else "") + m.group("c1") + row(m.group("r1"))
        if m.group("c2"):
            out += ":" + m.group("c2") + row(m.group("r2"))
        return out

    parts = re.split(r'("(?:[^"]|"")*")', f)
    return "".join(p if p.startswith('"') else _REF_ANY.sub(repl, p) for p in parts)


def _norm_f(f):
    return re.sub(r"\s+", "", f).replace("$", "").upper()
