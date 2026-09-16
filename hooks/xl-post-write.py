"""PostToolUse hook: after any Bash/Write/Edit that touched an .xlsx/.xlsm, run a quick xl lint
and hand the findings back to Claude as additionalContext. Silent when nothing applies.

Never blocks (PostToolUse cannot), never raises, bails out fast when no workbook path is in play.
Disable with env XL_HOOK_DISABLE=1.
"""
import json
import os
import re
import subprocess
import sys
import time

# where xl lives: XL_HOME if set, else %USERPROFILE%\tools\xl (what install.ps1 creates)
XLCLI = os.path.join(os.environ.get("XL_HOME") or os.path.join(os.path.expanduser("~"), "tools", "xl"), "xlcli.py")
RECENT_SECONDS = 240
MAX_MB = 30
# a cold kernel needs up to ~20 s to spawn plus ~1.2 s/MB to load; 45 s used to time out on the
# 30 MB files this hook accepts and then report a misleading IndexError
LINT_TIMEOUT = 100
PATH_RX = re.compile(r"""(?:"([^"\n]+?\.xls[xm])"|'([^'\n]+?\.xls[xm])'|([^\s"'<>|;&]+\.xls[xm]))""", re.I)


def out(ctx=None):
    if ctx:
        print(json.dumps({"hookSpecificOutput": {"hookEventName": "PostToolUse", "additionalContext": ctx}}))
    sys.exit(0)


def norm(p):
    p = p.strip()
    m = re.match(r"^/([A-Za-z])/(.*)$", p)
    if m:
        p = f"{m.group(1).upper()}:/{m.group(2)}"
    return os.path.normpath(p)


def main():
    if os.environ.get("XL_HOOK_DISABLE") == "1":
        out()
    try:
        data = json.load(sys.stdin)
    except Exception:
        out()
    tool = data.get("tool_name", "")
    ti = data.get("tool_input") or {}
    cands = []
    if tool in ("Write", "Edit"):
        fp = ti.get("file_path") or ""
        if fp.lower().endswith((".xlsx", ".xlsm")):
            cands.append(fp)
    elif tool == "Bash":
        cmd = ti.get("command") or ""
        if not re.search(r"\.xls[xm]\b", cmd, re.I):
            out()
        # an xl read/verify command is not a write; only `--save` or `exec` can change a file
        if re.match(r"^\s*(xl|py -3\.14 .*xlcli\.py)\b", cmd) and "--save" not in cmd and " exec" not in cmd:
            out()
        for m in PATH_RX.finditer(cmd):
            cands.append(m.group(1) or m.group(2) or m.group(3))
    else:
        out()
    now = time.time()
    paths = []
    for c in cands:
        p = norm(c)
        if p in paths or not os.path.isfile(p):
            continue
        st = os.stat(p)
        if now - st.st_mtime > RECENT_SECONDS:
            continue  # read, not written
        if st.st_size > MAX_MB * 1e6:
            paths.append((p, "too-big"))
            continue
        paths.append((p, "lint"))
    if not paths:
        out()
    ctx_parts = []
    for p, mode in paths[:2]:
        name = os.path.basename(p)
        if mode == "too-big":
            ctx_parts.append(f"[xl] {name} was just written ({os.path.getsize(p) / 1e6:.0f} MB). Run `xl lint` and `xl calc` on it before reporting done.")
            continue
        try:
            r = subprocess.run([sys.executable, XLCLI, "-s", "hook", "-t", str(LINT_TIMEOUT), "--json", "lint", p, "--quick"],
                               capture_output=True, text=True, timeout=LINT_TIMEOUT + 15, encoding="utf-8", errors="replace")
            lines = [ln for ln in r.stdout.strip().splitlines() if ln.startswith("{")]
            if not lines:
                why = (r.stderr or "").strip()[-300:] or f"no output (exit {r.returncode}); the kernel may still be loading the file"
                raise RuntimeError(why)
            d = json.loads(lines[-1])
        except subprocess.TimeoutExpired:
            ctx_parts.append(f"[xl] {name} was just written; quick lint timed out after {LINT_TIMEOUT}s (large file or cold kernel). Run `xl lint \"{p}\"` and `xl calc \"{p}\"` before reporting done.")
            continue
        except Exception as e:  # noqa: BLE001
            ctx_parts.append(f"[xl] {name} was just written; quick lint could not run ({type(e).__name__}: {str(e)[:200]}). Run `xl lint \"{p}\"` and `xl calc \"{p}\"` before reporting done.")
            continue
        sm = d.get("summary", {})
        head = f"[xl quick lint] {name}: errors={sm.get('errors', 0)} warnings={sm.get('warnings', 0)}"
        lines = [head]
        for f in d.get("findings", [])[:8]:
            if f["level"] == "info":
                continue
            lines.append(f"  {f['level']} {f['rule']} {f['where']}: {f['msg'][:110]}")
        lines.append(f"  Next: `xl calc \"{p}\"` for a live recalculation and the full error list; cite Sheet!Address for any figure you report.")
        ctx_parts.append("\n".join(lines))
    out("\n".join(ctx_parts)[:6000])


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        sys.exit(0)
