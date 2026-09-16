"""Cached values and calculation settings at the package level, without Excel.

inject(): write the correct cached value into formula cells, so a freshly built block (a
think-cell feed, a bridge) reads right the moment the file opens, while the workbook itself is
never recalculated and the formulas stay live. Ported from the think-cell build tooling of
2026-08-24, where each of the traps below produced a file Excel opened "Repaired" and read-only:

  1. An uncalculated formula returning "" is stored as a SELF-CLOSING <v/>. Testing for the
     literal <v> misses it and appends a second value element (schema-invalid).
  2. t="..." must be stripped from the <c> opening tag only; a whole-cell regex also eats
     <f t="shared"> and breaks every cell that shares that formula.
  3. A verifier that reads the value back with a loose <v>(.*)</v> finds whichever <v> comes
     first and passes a broken file. Every patched cell is asserted to hold exactly one <v>.

Also fixed on the way in: a self-closing <c .../> is no longer swallowed together with the next
cell (the old pattern could land a value in the neighbour), and workbook/rels parts are parsed
as XML, so files written by openpyxl (attribute order, absolute rels targets) resolve too.

calcpr(): read or set <calcPr> in xl/workbook.xml. A save through a Workbooks.Add() seed can
drop iterate="1", and fullCalcOnLoad="1" makes Excel recalculate on open, which silently
replaces every injected value. Only cells that already carry an <f> are patched; every other
zip part is copied across byte for byte.
"""
import math
import os
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .paths import backup_to_work, guard_write, under_sync_root

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

# <c ...> with quoted attribute values; group 1 = attributes, group 2 = "/>" or ">...</c>"
CELL_RX = re.compile(rb'<c\b((?:[^>/"]|"[^"]*"|/(?!>))*)(/>|>(.*?)</c>)', re.S)
REF_RX = re.compile(rb'\sr="([A-Za-z]+[0-9]+)"')
F_RX = re.compile(rb"<f\b(?:[^>/\"]|\"[^\"]*\"|/(?!>))*(?:/>|>.*?</f>)", re.S)
V_ANY = re.compile(rb"<v\b[^>]*/>|<v\b[^>]*>.*?</v>", re.S)
V_COUNT = re.compile(rb"<v[\s/>]")
T_ATTR = re.compile(rb'\st="[^"]*"')

# CT_Workbook children that may precede <calcPr>, in schema order
_BEFORE_CALCPR = ("definedNames", "externalReferences", "functionGroups", "sheets")


def _unesc(s):
    return s.replace("&quot;", '"').replace("&apos;", "'").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def sheet_part(z, sheet):
    """Zip part name of a worksheet, by name (exact, then case/space-insensitive, then unique substring)."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{NS_PKG}}}Relationship")}
    names = {}
    for s in wb.iter(f"{{{NS_MAIN}}}sheet"):
        t = targets.get(s.get(f"{{{NS_REL}}}id"))
        if t:
            names[s.get("name")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath("xl/" + t)
    if sheet in names:
        return names[sheet]
    want = sheet.strip().lower()
    for k, v in names.items():
        if k.strip().lower() == want:
            return v
    cands = [k for k in names if want in k.strip().lower()]
    if len(cands) == 1:
        return names[cands[0]]
    raise KeyError(f"sheet {sheet!r} not found; sheets: {list(names)}")


def _encode(val):
    """(t attribute or None, <v> text) for a cached value; raises on what a formula cell cannot cache."""
    if isinstance(val, bool):
        return b"b", b"1" if val else b"0"
    if isinstance(val, (int, float)):
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            raise ValueError(f"cannot cache {val!r}")
        return None, repr(f).encode()
    if isinstance(val, str):
        return b"str", escape(val).encode("utf-8")
    raise TypeError(f"unsupported cached value type {type(val).__name__}: {val!r}")


def _patch_cell(ref, attrs, inner, val):
    t, vtxt = _encode(val)
    attrs = T_ATTR.sub(b"", attrs, count=1)  # the <c> opening tag only (trap 2)
    if t:
        attrs += b' t="' + t + b'"'
    vel = b"<v>" + vtxt + b"</v>"
    if V_ANY.search(inner):  # covers the self-closing <v/> (trap 1)
        inner = V_ANY.sub(lambda _m: vel, inner, count=1)
    else:
        m = F_RX.search(inner)
        inner = inner[:m.end()] + vel + inner[m.end():]
    n = len(V_COUNT.findall(inner))
    if n != 1:  # trap 3
        raise RuntimeError(f"cell {ref} would hold {n} <v> elements: {inner[:200]!r}")
    return b"<c" + attrs + b">" + inner + b"</c>"


def _rewrite(src, dst, replace):
    """Copy every zip part of src into dst, substituting the parts in `replace` ({name: bytes})."""
    tmp = dst + ".xltmp"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replace.get(item.filename)
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)  # fresh header: no stale flags/extras
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, zin.read(item.filename) if data is None else data)
    os.replace(tmp, dst)


def _resolve_out(path, out, force):
    """In-place writes are allowed on a work copy only (work copy, or force=True plus a backup); a separate output
    must not land under a sync root unless force=True."""
    if out is None or os.path.abspath(out) == os.path.abspath(path):
        guard_write(path, force=force)
        return path, backup_to_work(path)
    if under_sync_root(out) and not force:
        raise PermissionError(f"{out} is under a OneDrive/Teams sync root. Write to the work dir and "
                              f"deliver through the gate, or pass force=True.")
    return out, None


def read_calcpr(path):
    """Attributes of <calcPr> as a dict ({} if the element is missing)."""
    with zipfile.ZipFile(path) as z:
        el = ET.fromstring(z.read("xl/workbook.xml")).find(f"{{{NS_MAIN}}}calcPr")
    return {} if el is None else dict(el.attrib)


def _calcpr_warning(attrs):
    if attrs.get("fullCalcOnLoad") in ("1", "true"):
        return "fullCalcOnLoad is set: Excel recalculates on open and replaces the injected values"
    if attrs.get("calcMode", "auto") != "manual":
        return ("calcMode is automatic: Excel may recalculate on open (older calcId) and replace the "
                "injected values; calcpr(path, set={'calcMode': 'manual'}) if the file must not recalculate")
    return None


def inject(path, sheet, values, out=None, force=False):
    """values: {"C570": 123.4, "D12": "FY25", "E3": True}. Returns a dict with the cells patched,
    skipped (no formula), missing (no such cell in the file), the backup, and a calcPr warning."""
    todo = {k.replace("$", "").upper(): v for k, v in values.items()}
    for k, v in todo.items():
        _encode(v)  # fail before anything is written
    with zipfile.ZipFile(path) as z:
        part = sheet_part(z, sheet)
        xml = z.read(part)
    head = xml[:4000]
    if re.search(rb"<\w+:worksheet\b", head):
        raise ValueError(f"{part} uses a namespace prefix on its elements; not supported")
    patched, no_formula, seen = [], [], set()

    def repl(m):
        attrs, tail, inner = m.group(1), m.group(2), m.group(3)
        rm = REF_RX.search(b" " + attrs)
        ref = rm.group(1).decode().upper() if rm else None
        if ref not in todo:
            return m.group(0)
        seen.add(ref)
        if tail == b"/>" or not F_RX.search(inner or b""):
            no_formula.append(ref)
            return m.group(0)
        patched.append(ref)
        return _patch_cell(ref, attrs, inner, todo[ref])

    new_xml = CELL_RX.sub(repl, xml)
    missing = sorted(set(todo) - seen)
    target, backup = _resolve_out(path, out, force)
    if patched:
        _rewrite(path, target, {part: new_xml})
    elif target != path:
        _rewrite(path, target, {})

    # read back what was written: exactly one <v> per patched cell, carrying the injected text
    with zipfile.ZipFile(target) as z:
        back = z.read(part)
    cells = {}
    for m in CELL_RX.finditer(back):
        rm = REF_RX.search(b" " + m.group(1))
        if rm and rm.group(1).decode().upper() in patched:
            cells[rm.group(1).decode().upper()] = m.group(3) or b""
    for ref in patched:
        inner = cells.get(ref, b"")
        vs = V_ANY.findall(inner)
        want = b"<v>" + _encode(todo[ref])[1] + b"</v>"
        if len(V_COUNT.findall(inner)) != 1 or vs[0] != want:
            raise RuntimeError(f"read-back failed on {ref}: {inner[:200]!r}")

    calc_attrs = read_calcpr(target)
    return {"file": str(target), "sheet": sheet, "part": part, "patched": patched,
            "no_formula": no_formula, "missing": missing, "backup": backup,
            "calcpr": calc_attrs, "warning": _calcpr_warning(calc_attrs)}


def calcpr(path, set=None, out=None, force=False):  # noqa: A002 - mirrors the CLI flag
    """Read <calcPr>, or with set={"calcMode": "manual", "iterate": "1", "fullCalcOnLoad": None}
    rewrite it (None removes an attribute; the element is created if missing)."""
    before = read_calcpr(path)
    if not set:
        return {"file": str(path), "calcpr": before, "warning": _calcpr_warning(before)}
    after = dict(before)
    for k, v in set.items():
        if v is None:
            after.pop(k, None)
        else:
            after[k] = str(v)
    tag = "<calcPr" + "".join(f' {k}="{escape(v, {chr(34): "&quot;"})}"' for k, v in after.items()) + "/>"
    with zipfile.ZipFile(path) as z:
        text = z.read("xl/workbook.xml").decode("utf-8")
    if re.search(r"<\w+:workbook\b", text[:2000]):
        raise ValueError("workbook.xml uses a namespace prefix; not supported")
    rx = re.compile(r"<calcPr\b(?:[^>/\"]|\"[^\"]*\"|/(?!>))*(?:/>|>.*?</calcPr>)", re.S)
    if rx.search(text):
        text = rx.sub(lambda _m: tag, text, count=1)
    else:
        for name in _BEFORE_CALCPR:
            m = None
            for m in re.finditer(rf"</{name}>|<{name}\b[^>]*/>", text):
                pass
            if m:
                text = text[:m.end()] + tag + text[m.end():]
                break
        else:
            raise ValueError("no <sheets> element in workbook.xml")
    target, backup = _resolve_out(path, out, force)
    _rewrite(path, target, {"xl/workbook.xml": text.encode("utf-8")})
    got = read_calcpr(target)
    if got != after:
        raise RuntimeError(f"calcPr read-back mismatch: {got} != {after}")
    return {"file": str(target), "before": before, "calcpr": got, "backup": backup,
            "warning": _calcpr_warning(got)}


def fmt_inject(res):
    lines = [f"INJECT {os.path.basename(res['file'])} [{res['sheet']}]  patched={len(res['patched'])}  "
             f"no_formula={len(res['no_formula'])}  missing={len(res['missing'])}"]
    if res["no_formula"]:
        lines.append("  skipped, no formula: " + ", ".join(res["no_formula"][:20]))
    if res["missing"]:
        lines.append("  not in the file: " + ", ".join(res["missing"][:20]))
    if res.get("backup"):
        lines.append(f"  backup: {res['backup']}")
    lines.append(f"  calcPr {res['calcpr']}")
    if res.get("warning"):
        lines.append(f"  WARNING {res['warning']}")
    return "\n".join(lines)


def fmt_calcpr(res):
    lines = [f"CALCPR {os.path.basename(res['file'])}  {res['calcpr']}"]
    if "before" in res:
        lines.append(f"  was {res['before']}")
    if res.get("backup"):
        lines.append(f"  backup: {res['backup']}")
    if res.get("warning"):
        lines.append(f"  WARNING {res['warning']}")
    return "\n".join(lines)
