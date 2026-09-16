"""Loading a feed spec (JSON), number-format presets, and the label lints."""
import json
import os
import re

from .builders import SpecError

# Number formats. "_sz" suppresses zeros (structural zeros only: a period the model does not
# populate yet). Never suppress zeros on a revenue or cost block: a real zero is information.
# Formats that may sit next to think-cell's literal "e" must end ";;@" or ";@".
FORMATS = {
    "number": "#,##0.0;(#,##0.0);0.0",
    "number0": "#,##0;(#,##0);0",
    "number_sz": "#,##0.0;(#,##0.0);;@",
    "number0_sz": "#,##0;(#,##0);;@",
    "pct": "0.0%;-0.0%;0.0%",
    "pct_sz": "0.0%;-0.0%;;@",
    "waterfall": "#,##0;(#,##0);;@",
    "waterfall1": "#,##0.0;(#,##0.0);;@",
    "count": "#,##0;-#,##0;0",
    "check": "#,##0.000;(#,##0.000);0.000",
    "year": "0",
    "text": "@",
}

TOP_KEYS = {"version", "workbook", "feed_sheet", "label_col", "first_col", "grids", "blocks",
            "options", "_comment", "description"}
OPTION_DEFAULTS = {
    "reader": "auto",          # auto | stream | xl
    "xl_session": "tcpipe",
    "run_dir": None,           # default: %LOCALAPPDATA%/xl/work/tcfeed/<stamp>_<stem>
    "calcpr": "restore",       # restore | restore+manual
    "check_tol": 1e-6,
    "deep_limit_mb": 25,       # compare every other sheet's cached values when the file is smaller
    "allow_errors": False,     # True: a computed Excel error is left uninjected and reported
    "reader_mb_limit": 8,      # auto reader: xl kernel below this size, streaming above
}

_PERIOD = re.compile(
    r"(19|20)\d{2}|\bFY\s?'?(\d{4}|\d{2})[A-Za-z]{0,2}\b|\b[1-4]Q\s?'?\d{2}\b|\bQ[1-4]\s?'?\d{2}\b"
    r"|\bH[12]\s?'?\d{2}\b|\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-\s']?\d{2,4}\b", re.I)


def fmt(code, where):
    if code is None:
        return None
    if not isinstance(code, str) or not code:
        raise SpecError(f"{where}: format must be a preset name or a format code")
    code = FORMATS.get(code, code)
    if code.strip().lower() in ("general", "estándar", "standard"):
        raise SpecError(f"{where}: never set 'General' (Excel rejects it on non-English locales); "
                        f"use an explicit format")
    return code


def states_period(label):
    return bool(_PERIOD.search(label or ""))


def lint_text(text, where, problems):
    if isinstance(text, str) and "%%" in text:
        problems.append(f"{where}: '%%' in {text[:60]!r} (a Python %-format that was never applied)")


def load(path):
    path = os.path.abspath(path)
    with open(path, encoding="utf-8") as fh:
        spec = json.load(fh)
    if not isinstance(spec, dict):
        raise SpecError("spec must be a JSON object")
    unknown = set(spec) - TOP_KEYS
    if unknown:
        raise SpecError(f"unknown top-level keys: {sorted(unknown)}")
    for k in ("workbook", "feed_sheet", "blocks"):
        if k not in spec:
            raise SpecError(f"spec needs '{k}'")
    wbp = os.path.expandvars(os.path.expanduser(spec["workbook"]))
    if not os.path.isabs(wbp):
        wbp = os.path.join(os.path.dirname(path), wbp)
    spec["workbook"] = os.path.normpath(wbp)
    opts = dict(OPTION_DEFAULTS)
    opts.update(spec.get("options") or {})
    bad = set(opts) - set(OPTION_DEFAULTS)
    if bad:
        raise SpecError(f"unknown options: {sorted(bad)}")
    if opts["reader"] not in ("auto", "stream", "xl"):
        raise SpecError("options.reader must be auto, stream or xl")
    if opts["calcpr"] not in ("restore", "restore+manual"):
        raise SpecError("options.calcpr must be 'restore' or 'restore+manual'")
    spec["options"] = opts
    spec["_path"] = path
    if not isinstance(spec["blocks"], list) or not spec["blocks"]:
        raise SpecError("'blocks' must be a non-empty list")
    return spec
