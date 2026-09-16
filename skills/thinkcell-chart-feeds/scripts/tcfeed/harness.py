"""COM build harness: a dedicated hidden Excel that writes a plan and never calculates.

Rules it enforces (each one has broken a real build):
  - DispatchEx, never an attach to the user's session; xl's add-in dialog guard on our PID.
  - Workbooks.Add() seed first, then Calculation = manual and CalculateBeforeSave = False
    (neither is settable before a workbook exists); both asserted before open, after open,
    before and after the save. CalculateBeforeSave fires even in manual mode.
  - Open with UpdateLinks=0; refuse a read-only open.
  - Nothing is ever calculated: no Calculate, no F9, no Range.Calculate.
  - Bulk Range reads (Value2 / Formula arrays), one array write per row.
  - Formats cloned from donor rows with Range.Copy(Destination) (no clipboard), then the
    contents cleared. Merged donor cells are refused.
  - Literal text goes in with a leading apostrophe whenever Excel could parse it (numbers,
    dates, percentages, TRUE/FALSE, a leading = + - @).
  - Colours are BGR through COM: Font.Color = 255 is RED. bgr("0000FF") is blue.
  - Worksheet.Names (the sheet-scoped names think-cell uses; Workbook.Names cannot see them)
    and the shape count are compared before the save; a mismatch aborts without saving.
  - Close without saving, close the seed, Quit, all in a finally; our PID is killed only if
    it outlives Quit and owns no visible window.
"""
import gc
import os
import re
import time
import traceback

from . import a1
from . import xlbridge
from . import xlsem as X

XL_MANUAL = -4135
XL_DOWN = -4121
XL_MOVE = 2
XL_RIGHT = -4152
XL_PREVIOUS = 2
XL_BY_ROWS = 1
XL_FORMULAS = -4123


class BuildError(RuntimeError):
    pass


def bgr(hexrgb):
    h = hexrgb.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return b * 65536 + g * 256 + r


BLUE = bgr("0000FF")
INPUT_FILL = bgr("FFFFCC")
GREY = bgr("595959")

_QUOTE_FIRST = set("=+-@#'\"(")


def literal(text):
    t = str(text)
    if (t[:1] in _QUOTE_FIRST or any(ch.isdigit() for ch in t) or "%" in t
            or t.strip().upper() in ("TRUE", "FALSE") or t != t.strip()):
        return "'" + t
    return t


def _grid(v):
    if not isinstance(v, tuple):
        return ((v,),)
    return tuple(r if isinstance(r, tuple) else (r,) for r in v)


class ExcelBuild:
    def __init__(self, path, sheet, log):
        self.path, self.sheet, self.log = os.path.abspath(path), sheet, log
        self.app = self.seed = self.wb = self.ws = None
        self.pid = None

    # ---------------------------------------------------------------- lifecycle
    def __enter__(self):
        try:
            self._start()
        except BaseException:
            self._teardown()
            raise
        return self

    def __exit__(self, exc_type, exc, tb):
        if tb is not None:
            # a Range object held by a frame of the traceback keeps Excel alive after Quit
            traceback.clear_frames(tb)
        self._teardown()
        return False

    def _start(self):
        import win32com.client as w32
        xc = xlbridge.excelcom()
        self.app = app = w32.DispatchEx("Excel.Application")
        self.pid = xc._pid_of_hwnd(app.Hwnd)
        xc.start_dialog_guard(self.pid)
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        for prop, val in (("AskToUpdateLinks", False), ("AutomationSecurity", 3)):
            try:
                setattr(app, prop, val)
            except Exception:  # noqa: BLE001
                pass
        self.seed = app.Workbooks.Add()
        app.Calculation = XL_MANUAL
        app.CalculateBeforeSave = False
        self._assert_calc("after the seed workbook")
        self.log(f"[com] hidden Excel pid {self.pid}: manual calculation, CalculateBeforeSave off")
        self.wb = app.Workbooks.Open(os.path.normpath(self.path), UpdateLinks=0, ReadOnly=False,
                                     IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        if self.wb.ReadOnly:
            raise BuildError("Excel opened the workbook read-only (open elsewhere, or a repaired package)")
        self._assert_calc("after opening the workbook")
        self.ws = xc._ws(self.wb, self.sheet)
        self.log(f"[com] opened {os.path.basename(self.path)}, sheet {self.ws.Name!r}")

    def _assert_calc(self, when):
        if self.app.Calculation != XL_MANUAL:
            raise BuildError(f"Calculation is not manual {when}")
        if self.app.CalculateBeforeSave:
            raise BuildError(f"CalculateBeforeSave is on {when}")

    def _teardown(self):
        steps = (("close workbook", lambda: self.wb is not None and self.wb.Close(SaveChanges=False)),
                 ("close seed", lambda: self.seed is not None and self.seed.Close(SaveChanges=False)),
                 ("quit", lambda: self.app is not None and self.app.Quit()))
        for label, fn in steps:
            try:
                fn()
            except Exception as e:  # noqa: BLE001 - every step must run
                self.log(f"[com] teardown: {label} failed: {e!r}"[:200])
        self.ws = self.wb = self.seed = self.app = None
        gc.collect()
        if not self.pid:
            return
        xc, pr = xlbridge.excelcom(), xlbridge.procs()
        t0 = time.time()
        while time.time() - t0 < 20 and pr.pid_alive(self.pid):
            time.sleep(0.25)
        if pr.pid_alive(self.pid):
            if any(w["visible"] and w["class"] == "XLMAIN" for w in xc.windows_of_pid(self.pid)):
                self.log(f"[com] pid {self.pid} has a visible window: NOT killed")
            else:
                self.log(f"[com] pid {self.pid} outlived Quit: killed={xc.kill_pid(self.pid)}")
        dismissed = xc.guard_log(self.pid)
        if dismissed:
            self.log(f"[com] dialogs dismissed on our instance: {dismissed}")
        xc.stop_dialog_guard(self.pid)
        self.log("[com] Excel closed")

    # ---------------------------------------------------------------- reads
    def rng(self, r0, c0, r1=None, c1=None):
        ws = self.ws
        return ws.Range(ws.Cells(r0, c0), ws.Cells(r1 or r0, c1 or c0))

    def read(self, r0, r1, c0, c1, prop="Value2"):
        return _grid(getattr(self.rng(r0, c0, r1, c1), prop))

    def cell(self, r, c):
        return self.read(r, r, c, c)[0][0]

    def last_used_row(self):
        hit = self.ws.Cells.Find("*", SearchOrder=XL_BY_ROWS, SearchDirection=XL_PREVIOUS, LookIn=XL_FORMULAS)
        return 0 if hit is None else int(hit.Row)

    def names(self):
        out = {}
        for n in self.ws.Names:
            out[str(n.Name).split("!", 1)[-1]] = {"refers_to": str(n.RefersTo), "visible": bool(n.Visible)}
        return out

    def shapes(self):
        return [(str(s.Name), int(s.TopLeftCell.Row), int(s.Placement)) for s in self.ws.Shapes]

    def cf_count(self):
        try:
            return int(self.ws.Cells.FormatConditions.Count)
        except Exception:  # noqa: BLE001
            return None

    # ---------------------------------------------------------------- writes
    def insert_rows(self, at, n):
        saved = []
        for s in self.ws.Shapes:
            saved.append((s.Name, s.Placement))
            if s.Placement != XL_MOVE:
                s.Placement = XL_MOVE          # move with the rows, never stretch across the insert
        try:
            self.ws.Rows(f"{at}:{at + n - 1}").Insert(Shift=XL_DOWN)
        finally:
            for name, pl in saved:
                try:
                    self.ws.Shapes(name).Placement = pl
                except Exception:  # noqa: BLE001
                    pass

    def assert_empty(self, r0, r1, c0, c1):
        f = self.read(r0, r1, c0, c1, "Formula")
        for i, row in enumerate(f):
            for j, v in enumerate(row):
                if v not in (None, ""):
                    raise BuildError(f"target cell {a1.addr(c0 + j, r0 + i)} is not empty ({str(v)[:40]!r})")

    def clone_format(self, src, dst, c0, c1):
        s = self.rng(src, c0, src, c1)
        if s.MergeCells is not False:
            raise BuildError(f"donor row {src} contains merged cells; never merge, pick another donor")
        d = self.rng(dst, c0, dst, c1)
        s.Copy(Destination=d)
        d.ClearContents()
        try:
            d.ClearComments()
        except Exception:  # noqa: BLE001
            pass

    def clear_formats(self, r, c0, c1):
        self.rng(r, c0, r, c1).ClearFormats()

    def write_row(self, r, c0, values):
        """values[k] goes to column c0+k: None (blank), '=formula', text, or a number."""
        arr = []
        for v in values:
            if v is None:
                arr.append(None)
            elif isinstance(v, str):
                arr.append(v if v.startswith("=") else literal(v))
            else:
                arr.append(float(v))
        long_text = any(isinstance(v, str) and len(v) > 250 for v in arr)
        if long_text:
            for k, v in enumerate(arr):
                if v is not None:
                    self.rng(r, c0 + k).Formula = v
        else:
            self.rng(r, c0, r, c0 + len(arr) - 1).Formula = (tuple(arr),)
        back_f = self.read(r, r, c0, c0 + len(arr) - 1, "Formula")[0]
        back_v = self.read(r, r, c0, c0 + len(arr) - 1, "Value2")[0]
        for k, v in enumerate(values):
            addr = a1.addr(c0 + k, r)
            got_f, got_v = back_f[k], back_v[k]
            if v is None:
                if got_f not in (None, ""):
                    raise BuildError(f"{addr} should be blank, holds {got_f!r}")
            elif isinstance(v, str) and v.startswith("="):
                if not (isinstance(got_f, str) and got_f.startswith("=")):
                    raise BuildError(f"{addr}: formula did not land ({got_f!r}); was the leading '=' lost?")
            elif isinstance(v, str):
                if got_v != v:
                    raise BuildError(f"{addr}: text read back as {got_v!r}, wrote {v!r}")
            elif not (X.is_num(got_v) and float(got_v) == float(v)):
                raise BuildError(f"{addr}: number read back as {got_v!r}, wrote {v!r}")
        return back_v

    def set_numfmt(self, r, c0, c1, code):
        rg = self.rng(r, c0, r, c1)
        rg.NumberFormat = code
        got = rg.NumberFormat
        return got if got != code else None

    def style(self, r, c0, c1, bold=None, italic=None, color=None, align=None):
        rg = self.rng(r, c0, r, c1)
        if bold is not None:
            rg.Font.Bold = bold
        if italic is not None:
            rg.Font.Italic = italic
        if color is not None:
            rg.Font.Color = color
        if align is not None:
            rg.HorizontalAlignment = align

    def style_input(self, r, c):
        rg = self.rng(r, c)
        rg.Font.Color = BLUE
        rg.Interior.Color = INPUT_FILL
        rg.BorderAround(1, 2)                   # xlContinuous, xlThin

    def save(self):
        self._assert_calc("before the save")
        self.wb.Save()
        self._assert_calc("after the save")
        self.log("[com] saved (manual calculation, CalculateBeforeSave off)")


DONOR_ALIAS = {"segment": ("segment", "data"), "inputs": ("inputs", "input"), "gap": ("blank",),
               "spacer": ("spacer", "blank"), "blank": ("blank",)}


def build(plan, computed, log):
    """Write the plan into the workbook. `computed`: {(row, col): value} (only used to record
    what Excel itself cached for the new formulas before injection)."""
    rep = {"insertions": [], "format_readback": [], "excel_cached_on_entry": {}}
    with ExcelBuild(plan.workbook, plan.feed_sheet, log) as xb:
        names0, shapes0, cf0 = xb.names(), xb.shapes(), xb.cf_count()
        rep.update(names_before=names0, shapes_before=len(shapes0), cf_before=cf0,
                   last_used_row_before=xb.last_used_row())
        # phase 1: every native insertion, in spec order
        for b in plan.blocks:
            if b.insert_time_row is None:
                continue
            t, n = b.insert_time_row, b.height
            got = xb.cell(t, b.anchor_col)
            if not (isinstance(got, str) and re.search(b.anchor_regex, got)):
                raise BuildError(f"{b.id}: row {t} holds {got!r}, expected the anchor {b.anchor_regex!r}")
            xb.insert_rows(t, n)
            moved = xb.cell(t + n, b.anchor_col)
            if moved != got:
                raise BuildError(f"{b.id}: after inserting {n} rows the anchor is not at row {t + n}")
            log(f"[build] {b.id}: inserted {n} rows at {t}; anchor now at row {t + n}")
            rep["insertions"].append({"block": b.id, "at": t, "rows": n})
        # phase 2: write every block at its final rows
        same = zero = other = 0
        for b in plan.blocks:
            if b.insert_time_row is None:
                xb.assert_empty(b.first_row, b.last_row, b.label_col, b.last_col)
            for pr in b.rows:
                r = b.row_of(pr.offset)
                donor = next((b.donors[k] for k in DONOR_ALIAS.get(pr.kind, (pr.kind,)) if k in b.donors), None)
                if donor is not None:
                    xb.clone_format(plan.to_new(donor), r, b.label_col, b.last_col)
                elif b.insert_time_row is not None:
                    xb.clear_formats(r, b.label_col, b.last_col)
                vals = plan.row_values(b, pr)
                cols = sorted(vals)
                if any(v is not None for v in vals.values()):
                    back = xb.write_row(r, cols[0], [vals[c] for c in cols])
                    for k, c in enumerate(cols):
                        cell = b.cells.get((pr.offset, c))
                        if cell is None or cell.kind != "formula":
                            continue
                        got, want = back[k], computed.get((r, c))
                        if X.same_value(got, want, rel=1e-9, abs_tol=1e-9):
                            same += 1
                        elif got in (None, "", 0, 0.0):
                            zero += 1
                            rep.setdefault("excel_cached_zero_cells", []).append(a1.addr(c, r))
                        else:
                            other += 1
                            rep.setdefault("excel_cached_other_cells", []).append(
                                f"{a1.addr(c, r)} excel={got!r} computed={want!r}")
                if pr.fmt and pr.kind not in ("title", "note", "header", "spacer", "blank", "gap"):
                    bad = xb.set_numfmt(r, b.first_col, b.last_col, pr.fmt)
                    if bad is not None:
                        rep["format_readback"].append({"row": r, "wrote": pr.fmt, "read": bad})
                if donor is None:
                    if pr.kind == "title":
                        xb.style(r, b.label_col, b.label_col, bold=True)
                    elif pr.kind == "note":
                        xb.style(r, b.label_col, b.label_col, italic=True, color=GREY)
                    elif pr.kind == "header":
                        xb.style(r, b.label_col, b.last_col, bold=True)
                        xb.style(r, b.first_col, b.last_col, align=XL_RIGHT)
                    elif pr.kind == "check":
                        xb.style(r, b.label_col, b.last_col, italic=True)
                for (off, c), cell in b.cells.items():
                    if off == pr.offset and cell.style == "input":
                        xb.style_input(r, c)
            log(f"[build] {b.id}: wrote rows {b.first_row}-{b.last_row}")
        rep["excel_cached_on_entry"] = {"equal_to_computed": same, "zero_or_blank": zero, "other": other}
        # gates before the save
        names1, shapes1, cf1 = xb.names(), xb.shapes(), xb.cf_count()
        rep.update(names_after=names1, shapes_after=len(shapes1), cf_after=cf1)
        if set(names1) != set(names0):
            raise BuildError(f"sheet-scoped names changed: added {sorted(set(names1) - set(names0))}, "
                             f"lost {sorted(set(names0) - set(names1))}; not saved")
        for k, v in names0.items():
            want = a1.shift_sheet_refs(v["refers_to"], plan.feed_sheet, plan.to_new)
            if a1.norm_ref_text(want) != a1.norm_ref_text(names1[k]["refers_to"]):
                raise BuildError(f"name {k}: refers to {names1[k]['refers_to']}, expected {want}; not saved")
            if "#REF!" in names1[k]["refers_to"]:
                raise BuildError(f"name {k} is broken ({names1[k]['refers_to']}); not saved")
        if len(shapes1) != len(shapes0):
            raise BuildError(f"shape count {len(shapes0)} -> {len(shapes1)}; not saved")
        if cf0 is not None and cf1 is not None and cf1 > cf0:
            log(f"[build] WARNING conditional-format rules {cf0} -> {cf1} (donor rows carry rules; "
                f"repeated format pastes pile them up and can hang Save)")
        xb.save()
    return rep
