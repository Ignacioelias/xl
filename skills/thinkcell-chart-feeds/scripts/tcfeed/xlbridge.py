"""The only door to the `xl` tool (skill xl-repl): imports from its package, calls its CLI.

XL_HOME (env) or %USERPROFILE%/tools/xl. Nothing here modifies xl.
"""
import json
import os
import subprocess
import sys

XL_HOME = os.environ.get("XL_HOME") or os.path.join(os.path.expanduser("~"), "tools", "xl")


def _ensure():
    if not os.path.isfile(os.path.join(XL_HOME, "xl", "cached.py")):
        raise RuntimeError(f"xl not found at {XL_HOME} (set XL_HOME); tcfeed needs xl.cached and xl.excelcom")
    if XL_HOME not in sys.path:
        sys.path.insert(0, XL_HOME)


def cached():
    _ensure()
    from xl import cached as m
    return m


def excelcom():
    _ensure()
    from xl import excelcom as m
    return m


def procs():
    _ensure()
    from xl import procs as m
    return m


def paths():
    _ensure()
    from xl import paths as m
    return m


def exec_code(code, session, timeout=900):
    _ensure()
    from xl.client import exec_code as run
    return run(code, session=session, timeout=timeout, max_chars=200_000_000)


def cli(args, session, timeout=900):
    """Run `xl -s SESSION -t TIMEOUT <args>` with this interpreter. Returns (exit, stdout, stderr)."""
    _ensure()
    cmd = [sys.executable, os.path.join(XL_HOME, "xlcli.py"), "-s", session, "-t", str(int(timeout))] + list(args)
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       timeout=timeout + 60)
    return p.returncode, p.stdout, p.stderr


def check_open(path, session, wait=120):
    code, out, err = cli(["check-open", os.path.abspath(path), "--wait", str(wait)], session, timeout=wait + 120)
    return code, (out + ("\n" + err if err.strip() else "")).strip()


def calc_verify(path, session, timeout=900):
    """`xl calc PATH --verify --json` (read-only open, full rebuild, nothing saved)."""
    code, out, err = cli(["calc", os.path.abspath(path), "--verify", "--json"], session, timeout=timeout)
    lines = [ln for ln in out.splitlines() if ln.startswith("{")]
    res = json.loads(lines[-1]) if lines else None
    return code, res, err.strip()


def stop(session):
    code, out, err = cli(["stop"], session, timeout=60)
    return (out + err).strip()
