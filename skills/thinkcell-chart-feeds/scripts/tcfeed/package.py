"""Package-level reads and the one package-level write tcfeed does itself (dropping
[trash] parts). Cached values and calcPr are xl's job (xl.cached)."""
import os
import re
import zipfile
from xml.etree import ElementTree as ET

from . import xlbridge

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def workbook_meta(path):
    """(sheet names in order, [(name, localSheetId or None, hidden, text)])."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = [s.get("name") for s in root.iter(NS + "sheet")]
    names = []
    dn = root.find(NS + "definedNames")
    if dn is not None:
        for d in dn.iter(NS + "definedName"):
            lid = d.get("localSheetId")
            names.append((d.get("name"), int(lid) if lid is not None else None,
                          d.get("hidden") in ("1", "true"), d.text or ""))
    return sheets, names


def local_names(path, sheet):
    """Names scoped to `sheet` (what Worksheet.Names returns; think-cell stores its links
    here, hidden). {name: refers-to text}."""
    sheets, names = workbook_meta(path)
    idx = [i for i, s in enumerate(sheets) if s.lower() == sheet.lower()]
    if not idx:
        raise KeyError(f"sheet {sheet!r} not in {sheets}")
    return {n: {"text": t, "hidden": h} for n, lid, h, t in names if lid == idx[0]}


def global_names(path):
    _, names = workbook_meta(path)
    return {n: t for n, lid, h, t in names if lid is None}


def parts(path):
    with zipfile.ZipFile(path) as z:
        return z.namelist()


def trash_parts(path):
    return [n for n in parts(path) if n.startswith("[trash]/")]


def drop_parts(path, drop, out):
    """Copy every part except `drop` into `out` (fresh zip headers, same compression)."""
    drop = set(drop)
    tmp = out + ".tcftmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop:
                continue
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, zin.read(item.filename))
    os.replace(tmp, out)


def sheet_xml(path, sheet):
    cached = xlbridge.cached()
    with zipfile.ZipFile(path) as z:
        return z.read(cached.sheet_part(z, sheet)).decode("utf-8", "replace")


def merged_ranges(path, sheet):
    return re.findall(r'<mergeCell\s+ref="([^"]+)"', sheet_xml(path, sheet))


def is_locked(path):
    """(True, why) when another process (a live Excel) holds the file."""
    folder, name = os.path.split(os.path.abspath(path))
    for owner in ("~$" + name, "~$" + name[2:]):
        if os.path.exists(os.path.join(folder, owner)):
            return True, f"owner file {owner} exists (the workbook is open somewhere)"
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        return True, "the file refuses a write handle (held open by another process)"
    return False, ""
