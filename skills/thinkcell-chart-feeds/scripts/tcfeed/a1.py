"""A1 reference helpers. Pure Python, no Excel."""
import re

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_PLAIN_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_LOOKS_A1 = re.compile(r"^[A-Za-z]{1,3}\d+$")
_LOOKS_R1C1 = re.compile(r"^[Rr]\d*[Cc]?\d*$|^[Cc]\d*$")


def col_letter(n):
    n = int(n)
    if not 1 <= n <= 16384:
        raise ValueError(f"column {n} out of range")
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(c):
    if isinstance(c, int):
        return c
    s = str(c).strip().replace("$", "").upper()
    if not re.fullmatch(r"[A-Z]{1,3}", s):
        raise ValueError(f"bad column {c!r}")
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def split_cell(ref):
    """'$C$5' / 'C5' -> (3, 5)."""
    m = _CELL.match(str(ref).strip())
    if not m:
        raise ValueError(f"bad cell reference {ref!r}")
    return col_index(m.group(1)), int(m.group(2))


def addr(col, row, absolute=False):
    c = col_letter(col_index(col))
    return f"${c}${row}" if absolute else f"{c}{row}"


def sheet_prefix(name):
    """'Sheet!' or quoted "'Sheet name'!" when Excel needs the quotes (spaces, a name that
    looks like a cell or R1C1 reference, a leading digit, ...)."""
    if (_PLAIN_SHEET.match(name) and not _LOOKS_A1.match(name) and not _LOOKS_R1C1.match(name)
            and name.upper() not in ("TRUE", "FALSE")):
        return name + "!"
    return "'" + name.replace("'", "''") + "'!"


def row_range(sheet, row, c0, c1):
    return f"{sheet_prefix(sheet)}${col_letter(c0)}${row}:${col_letter(c1)}${row}"


def rect_range(sheet, r0, r1, c0, c1):
    return f"{sheet_prefix(sheet)}${col_letter(c0)}${r0}:${col_letter(c1)}${r1}"


def col_range(sheet, col, r0, r1):
    c = col_letter(col_index(col))
    return f"{sheet_prefix(sheet)}${c}${r0}:${c}${r1}"


def cell_ref(sheet, col, row, absolute=True):
    return sheet_prefix(sheet) + addr(col, row, absolute)


_SHEET_REF = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Za-z]{1,3}\$?\d+)$")


def parse_cell_ref(ref):
    """"Sheet!$C$5" or "'Sheet name'!C5" -> (sheet, col, row)."""
    m = _SHEET_REF.match(str(ref).strip())
    if not m:
        raise ValueError(f"expected Sheet!A1, got {ref!r}")
    sheet = m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2)
    col, row = split_cell(m.group(3))
    return sheet, col, row


# a sheet-qualified reference or range inside a defined name or formula text
REF_TOKEN = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!"
    r"(?P<c1>\$?[A-Za-z]{1,3})(?P<r1>\$?\d+)(?::(?P<c2>\$?[A-Za-z]{1,3})(?P<r2>\$?\d+))?")


def unquote_sheet(token):
    if token.startswith("'") and token.endswith("'"):
        return token[1:-1].replace("''", "'")
    return token


def shift_sheet_refs(text, sheet, to_new):
    """Rewrite every reference to `sheet` in `text` with row -> to_new(row). This is what a
    native Rows.Insert does to formulas and defined names that point at the sheet."""
    def fix_row(tok):
        dollar = "$" if tok.startswith("$") else ""
        return dollar + str(to_new(int(tok.lstrip("$"))))

    def repl(m):
        if unquote_sheet(m.group("sheet")).lower() != sheet.lower():
            return m.group(0)
        out = f"{m.group('sheet')}!{m.group('c1')}{fix_row(m.group('r1'))}"
        if m.group("c2"):
            out += f":{m.group('c2')}{fix_row(m.group('r2'))}"
        return out

    return REF_TOKEN.sub(repl, text)


def norm_ref_text(text):
    """Compare defined-name text loosely: no '$', no quotes around plain sheet names, no '='."""
    t = text.strip().lstrip("=")
    t = REF_TOKEN.sub(lambda m: f"{unquote_sheet(m.group('sheet')).lower()}!{m.group(0).split('!', 1)[1]}", t)
    return t.replace("$", "").upper()
