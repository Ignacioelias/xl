"""PowerPoint lint and render through COM.

pptx-toolkit rules baked in: DispatchEx attaches to the USER's PowerPoint (unlike Excel), so
never Quit, never touch Visible/WindowState, always open a COPY with WithWindow=False and always
close the presentation, otherwise the file stays locked headlessly.
"""
import os
import shutil
import time
from pathlib import Path

from .paths import WORK

MSO_TRUE = -1
PP_PLACEHOLDER = 14
MSO_PICTURE = 13
MSO_CHART = 3


def _app():
    import pythoncom
    import win32com.client as w32
    pythoncom.CoInitialize()
    return w32.DispatchEx("PowerPoint.Application")


def _open_copy(path):
    src = Path(path)
    dst = WORK / f"pplint_{int(time.time() * 1000)}_{src.name}"
    shutil.copyfile(src, dst)
    app = _app()
    pres = app.Presentations.Open(str(dst), True, False, False)  # ReadOnly, Untitled, WithWindow
    return app, pres, dst


def _rect(sh):
    return (sh.Left, sh.Top, sh.Left + sh.Width, sh.Top + sh.Height)


def _overlap_frac(a, b):
    x1, y1 = max(a[0], b[0]), max(a[1], b[1])
    x2, y2 = min(a[2], b[2]), min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    area_a = max((a[2] - a[0]) * (a[3] - a[1]), 1e-6)
    return (x2 - x1) * (y2 - y1) / area_a


def lint(path, slides=None, min_font=8.0, cap_per_rule=40):
    """Rules: off-slide, text-overflow, text-occluded, empty-placeholder, tiny-font, empty-chart.
    (picture-stretched and notes-missing were stubs that never emitted; removed 2026-09-09, and
    the stub's ScaleHeight call was a mutation on the linted copy.)"""
    app, pres, tmp = _open_copy(path)
    findings, counts = [], {}

    def add(level, rule, where, msg):
        counts[rule] = counts.get(rule, 0) + 1
        if counts[rule] <= cap_per_rule:
            findings.append({"level": level, "rule": rule, "where": where, "msg": msg})

    try:
        W, H = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        for sl in pres.Slides:
            i = sl.SlideIndex
            if slides and i not in slides:
                continue
            shapes = [sl.Shapes.Item(k) for k in range(1, sl.Shapes.Count + 1)]
            info = []
            for z, sh in enumerate(shapes):
                try:
                    if int(sh.Visible) == 0:
                        continue  # hidden shapes neither cover nor count
                    r = _rect(sh)
                    has_text = bool(sh.HasTextFrame) and bool(sh.TextFrame.HasText)
                    text = sh.TextFrame.TextRange.Text.strip() if has_text else ""
                    opaque = False
                    try:
                        opaque = bool(sh.Fill.Visible) and float(sh.Fill.Transparency) < 0.5 and sh.Type != 17
                    except Exception:
                        opaque = False
                    if sh.Type in (MSO_PICTURE,):
                        opaque = True
                    info.append({"z": z, "sh": sh, "rect": r, "text": text, "has_text": has_text, "opaque": opaque, "name": sh.Name})
                except Exception:
                    continue
            for d in info:
                sh, r, where = d["sh"], d["rect"], f"slide {i} '{d['name'][:28]}'"
                if r[0] < -2 or r[1] < -2 or r[2] > W + 2 or r[3] > H + 2:
                    add("warn", "off-slide", where, f"extends beyond the slide ({r[0]:.0f},{r[1]:.0f})-({r[2]:.0f},{r[3]:.0f}) vs {W:.0f}x{H:.0f}")
                if d["has_text"]:
                    try:
                        tr = sh.TextFrame.TextRange
                        rotated = abs(float(sh.Rotation)) > 1 or int(sh.TextFrame.Orientation) != 1
                        shrink = False
                        try:
                            shrink = int(sh.TextFrame2.AutoSize) == 2  # msoAutoSizeTextToFitShape
                        except Exception:
                            pass
                        # 20% + 4pt tolerance: BoundHeight ignores margins and small font substitutions;
                        # rotated / vertical text reports a misleading height (validated on slide 25)
                        if not rotated and not shrink and sh.TextFrame.AutoSize == 0 \
                                and tr.BoundHeight > sh.Height * 1.2 + 4:
                            add("warn", "text-overflow", where, f"text height {tr.BoundHeight:.0f} > shape {sh.Height:.0f}: {d['text'][:50]!r}")
                        if not rotated and sh.TextFrame.WordWrap == 0 and tr.BoundWidth > sh.Width * 1.2 + 4:
                            add("warn", "text-overflow", where, f"text width {tr.BoundWidth:.0f} > shape {sh.Width:.0f}")
                        small = [tr.Runs(k).Font.Size for k in range(1, min(tr.Runs().Count, 40) + 1)]
                        small = [s for s in small if s and s < min_font]
                        if small:
                            add("warn", "tiny-font", where, f"{len(small)} runs below {min_font}pt (min {min(small):.0f}pt)")
                    except Exception:
                        pass
                    cover = 0.0
                    for e in info:
                        if e["z"] > d["z"] and e["opaque"] and not e["has_text"]:
                            cover = max(cover, _overlap_frac(r, e["rect"]))
                    if cover >= 0.5:
                        add("warn", "text-occluded", where, f"{cover:.0%} covered by a later opaque shape: {d['text'][:50]!r}")
                try:
                    if sh.Type == PP_PLACEHOLDER and not d["has_text"] and not sh.HasChart and not sh.HasTable:
                        add("info", "empty-placeholder", where, "placeholder with no content")
                except Exception:
                    pass
                try:
                    if sh.HasChart and sh.Chart.SeriesCollection().Count == 0:
                        add("error", "empty-chart", where, "chart with no series")
                except Exception:
                    pass
        summary = {"errors": sum(1 for f in findings if f["level"] == "error"),
                   "warnings": sum(1 for f in findings if f["level"] == "warn"),
                   "infos": sum(1 for f in findings if f["level"] == "info"), "by_rule": counts,
                   "slides": pres.Slides.Count}
        return {"file": str(path), "summary": summary, "findings": findings}
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            os.unlink(tmp)
        except OSError:
            pass


def render(path, slide=None, out_dir=None, width=1600, height=900):
    app, pres, tmp = _open_copy(path)
    out_dir = Path(out_dir or WORK)
    out_dir.mkdir(parents=True, exist_ok=True)
    outs = []
    try:
        idxs = [slide] if slide else range(1, pres.Slides.Count + 1)
        for i in idxs:
            p = out_dir / f"{Path(path).stem}_s{i:03d}.png"
            pres.Slides.Item(i).Export(str(p), "PNG", width, height)
            size = os.path.getsize(p) if p.exists() else 0
            if size == 0:
                raise RuntimeError(f"slide {i} exported 0 bytes")
            outs.append({"slide": i, "png": str(p), "bytes": size})
        return outs
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            os.unlink(tmp)
        except OSError:
            pass


def fmt_lint(res):
    s = res["summary"]
    lines = [f"PPTX LINT {os.path.basename(res['file'])}  slides={s['slides']}  errors={s['errors']} warnings={s['warnings']} infos={s['infos']}"]
    order = {"error": 0, "warn": 1, "info": 2}
    for f in sorted(res["findings"], key=lambda f: (order[f["level"]], f["rule"], f["where"])):
        lines.append(f"  {f['level']:<5} {f['rule']:<18} {f['where']:<40} {f['msg']}")
    return "\n".join(lines)
