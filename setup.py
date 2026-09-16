"""xl installer / uninstaller, run by install.ps1 (or directly: py -3.14 setup.py install).

Installs from the folder this file sits in, which is either a clone of the repository
(https://github.com/Ignacioelias/xl) or a shared-folder distribution with a payload/ folder:
  the package          -> %USERPROFILE%\\tools\\xl        (XL_HOME overrides; a git clone there is used as is)
  skills/<name>        -> %USERPROFILE%\\.claude\\skills\\<name>   (xl-repl, thinkcell-chart-feeds)
  hooks/               -> %USERPROFILE%\\.claude\\hooks\\xl-post-write.{js,py}
  shims                -> %USERPROFILE%\\bin\\xl, xl.cmd   (+ %USERPROFILE%\\bin on the user PATH)
  settings             -> one PostToolUse entry in %USERPROFILE%\\.claude\\settings.json
Then runs the checks: package imports, Excel COM, kernel start, lint fixture, clean open,
cached-value injection (think-cell feeds).

Everything is idempotent: run it again to upgrade. `uninstall` reverses all of it, except the
Python packages and Python itself.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
PAYLOAD = os.path.join(HERE, "payload")
if os.path.isdir(PAYLOAD):  # shared-folder distribution
    SRC_PKG = os.path.join(PAYLOAD, "xl")
    SRC_SKILLS = os.path.join(PAYLOAD, "skill")
    SRC_HOOKS = os.path.join(PAYLOAD, "hooks")
else:  # a clone (or zip) of the repository
    SRC_PKG = HERE
    SRC_SKILLS = os.path.join(HERE, "skills")
    SRC_HOOKS = os.path.join(HERE, "hooks")
ISSUES = "https://github.com/Ignacioelias/xl/issues"
HOME = os.path.expanduser("~")
XL_HOME = os.environ.get("XL_HOME") or os.path.join(HOME, "tools", "xl")
CLAUDE = os.path.join(HOME, ".claude")
BIN = os.path.join(HOME, "bin")
HOOK_MATCHER = "Bash|Write|Edit"
HOOK_MARK = "xl-post-write"
SKILLS = ("xl-repl", "thinkcell-chart-feeds")


def say(msg: str) -> None:
    print(msg, flush=True)


def step(n: int, msg: str) -> None:
    say(f"\n[{n}/7] {msg}")


# ----------------------------------------------------------------------------- copy helpers
SKIP_NAMES = {"__pycache__", ".git", ".playwright-mcp"}
SKIP_SUFFIXES = (".pyc", ".log")


def _wanted(name: str) -> bool:
    return name not in SKIP_NAMES and not name.endswith(SKIP_SUFFIXES)


def copytree(src: str, dst: str) -> None:
    """Sync src into dst file by file. Never removes dst itself; removes only files that exist
    in dst but not in src, and never touches .git or caches (so a development checkout survives)."""
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if _wanted(d)]
        rel = os.path.relpath(root, src)
        droot = dst if rel == "." else os.path.join(dst, rel)
        os.makedirs(droot, exist_ok=True)
        for f in files:
            if _wanted(f):
                target = os.path.join(droot, f)
                if os.path.isfile(target):
                    os.chmod(target, 0o666)  # clear read-only (SharePoint copies are sometimes RO)
                shutil.copy2(os.path.join(root, f), target)
    # prune what the payload no longer has
    for root, dirs, files in os.walk(dst, topdown=False):
        dirs[:] = [d for d in dirs if _wanted(d)]
        rel = os.path.relpath(root, dst)
        sroot = src if rel == "." else os.path.join(src, rel)
        for f in files:
            if _wanted(f) and not os.path.isfile(os.path.join(sroot, f)):
                os.chmod(os.path.join(root, f), 0o666)
                os.remove(os.path.join(root, f))
        for d in dirs:
            p = os.path.join(root, d)
            if not os.path.isdir(os.path.join(sroot, d)) and not os.listdir(p):
                os.rmdir(p)


def is_dev_checkout(path: str) -> bool:
    return os.path.isdir(os.path.join(path, ".git"))


def write(path: str, text: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


# ----------------------------------------------------------------------------- settings.json
def load_settings() -> dict:
    p = os.path.join(CLAUDE, "settings.json")
    if not os.path.isfile(p):
        return {}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_settings(d: dict) -> None:
    p = os.path.join(CLAUDE, "settings.json")
    os.makedirs(CLAUDE, exist_ok=True)
    if os.path.isfile(p):
        shutil.copy2(p, p + f".bak-{time.strftime('%Y%m%d-%H%M%S')}")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2)
        f.write("\n")


def hook_command() -> str:
    node = shutil.which("node")
    if node:
        js = os.path.join(CLAUDE, "hooks", "xl-post-write.js").replace("\\", "/")
        return f'"{node.replace(chr(92), "/")}" "{js}"'
    py = os.path.join(CLAUDE, "hooks", "xl-post-write.py").replace("\\", "/")
    return f'py -3.14 "{py}"'


def register_hook() -> None:
    d = load_settings()
    hooks = d.setdefault("hooks", {})
    post = hooks.setdefault("PostToolUse", [])
    # drop any previous xl entry, then add ours
    post[:] = [e for e in post if not any(HOOK_MARK in h.get("command", "") for h in e.get("hooks", []))]
    post.append({
        "matcher": HOOK_MATCHER,
        "hooks": [{
            "type": "command",
            "command": hook_command(),
            "timeout": 75,
            "statusMessage": "xl quick lint on the workbook just written",
        }],
    })
    save_settings(d)


def unregister_hook() -> None:
    d = load_settings()
    post = d.get("hooks", {}).get("PostToolUse")
    if not post:
        return
    post[:] = [e for e in post if not any(HOOK_MARK in h.get("command", "") for h in e.get("hooks", []))]
    save_settings(d)


# ----------------------------------------------------------------------------- PATH
def add_to_user_path(folder: str) -> bool:
    """Append folder to the user's PATH (registry) if missing. Returns True if it changed."""
    if os.environ.get("XL_SETUP_NO_PATH") == "1":  # test runs against a scratch home
        return False
    try:
        import winreg
    except ImportError:
        return False
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ | winreg.KEY_WRITE) as k:
        try:
            cur, kind = winreg.QueryValueEx(k, "Path")
        except FileNotFoundError:
            cur, kind = "", winreg.REG_EXPAND_SZ
        parts = [p for p in cur.split(";") if p]
        if any(os.path.normcase(p.rstrip("\\")) == os.path.normcase(folder) for p in parts):
            return False
        parts.append(folder)
        winreg.SetValueEx(k, "Path", 0, kind, ";".join(parts))
    # tell running shells (Explorer) the environment changed
    try:
        import ctypes
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x1A, 0, "Environment", 2, 5000, None)
    except Exception:  # noqa: BLE001
        pass
    return True


# ----------------------------------------------------------------------------- checks
def run(cmd: list[str], timeout: int = 120) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, encoding="utf-8", errors="replace")


def check(label: str, ok: bool, detail: str = "") -> bool:
    say(f"   {'PASS' if ok else 'FAIL'}  {label}{('  ' + detail) if detail else ''}")
    return ok


def doctor() -> bool:
    xlcli = os.path.join(XL_HOME, "xlcli.py")
    py = [sys.executable]
    results = []

    r = run(py + ["-c", "import openpyxl, win32com.client, PIL; print(openpyxl.__version__)"])
    results.append(check("Python packages (openpyxl, pywin32, pillow)", r.returncode == 0, r.stdout.strip() or r.stderr.strip()[-200:]))

    r = run(py + ["-c", "import win32com.client as w; x=w.DispatchEx('Excel.Application'); x.Visible=False; v=x.Version; x.Quit(); print(v)"], timeout=90)
    results.append(check("Excel reachable through COM", r.returncode == 0, ("Excel " + r.stdout.strip()) if r.returncode == 0 else r.stderr.strip()[-200:]))

    r = run(py + [xlcli, "-s", "setup", "exec", "-c", "print('kernel ok')"], timeout=120)
    results.append(check("xl kernel starts and answers", "kernel ok" in r.stdout, "" if "kernel ok" in r.stdout else (r.stdout + r.stderr).strip()[-200:]))

    fixture = os.path.join(os.environ.get("LOCALAPPDATA", HOME), "xl", "work", "fixture_defects.xlsx")
    os.makedirs(os.path.dirname(fixture), exist_ok=True)
    r = run(py + [os.path.join(XL_HOME, "bench", "make_fixture.py"), fixture, "--check"], timeout=300)
    results.append(check("Lint fixture: every rule fires on its planted defect", r.returncode == 0, "" if r.returncode == 0 else (r.stdout + r.stderr).strip()[-300:]))

    r = run(py + [xlcli, "-s", "setup", "check-open", fixture], timeout=180)
    results.append(check("Excel opens the fixture without a repair prompt", r.returncode == 0, "" if r.returncode == 0 else (r.stdout + r.stderr).strip()[-200:]))

    run(py + [xlcli, "-s", "setup", "stop"], timeout=60)

    r = run(py + [os.path.join(XL_HOME, "bench", "inject_check.py"), os.path.join(os.path.dirname(fixture), "inject_check")], timeout=300)
    results.append(check("Think-cell feeds: injected values read by Excel, no repair", r.returncode == 0, "" if r.returncode == 0 else (r.stdout + r.stderr).strip()[-300:]))

    ok = all(results)
    say("\n   " + ("All checks passed." if ok else f"One or more checks failed. Please open an issue with this output: {ISSUES}"))
    return ok


# ----------------------------------------------------------------------------- install / uninstall
def install() -> int:
    say(f"Installing xl for {os.environ.get('USERNAME', HOME)}")
    say(f"  package  -> {XL_HOME}")
    say(f"  skills   -> {os.path.join(CLAUDE, 'skills')}\\{{{','.join(SKILLS)}}}")
    say(f"  hook     -> {os.path.join(CLAUDE, 'hooks')}")
    say(f"  command  -> {os.path.join(BIN, 'xl')}")

    step(1, "Python packages")
    r = run([sys.executable, "-m", "pip", "install", "--user", "--quiet", "--upgrade", "openpyxl>=3.1,<4", "pywin32>=306", "pillow>=10"], timeout=600)
    if r.returncode != 0:
        say(r.stderr[-800:])
        say("   pip failed. If you are behind a corporate proxy, run install.ps1 again from a terminal where 'pip install' works.")
        return 1
    say("   ok")

    step(2, "Package")
    same = os.path.normcase(os.path.abspath(SRC_PKG)) == os.path.normcase(os.path.abspath(XL_HOME))
    if same or is_dev_checkout(XL_HOME):
        say(f"   {XL_HOME} is a git checkout or the install source: used as is (update it with git pull)")
    else:
        copytree(SRC_PKG, XL_HOME)
        say(f"   {XL_HOME}")

    step(3, "Command shims")
    xlcli = os.path.join(XL_HOME, "xlcli.py")
    write(os.path.join(BIN, "xl"), f'#!/bin/sh\nexec py -3.14 "{xlcli.replace(chr(92), "/")}" "$@"\n')
    write(os.path.join(BIN, "xl.cmd"), f'@echo off\r\npy -3.14 "{xlcli}" %*\r\n')
    changed = add_to_user_path(BIN)
    say(f"   {BIN} {'added to your PATH' if changed else 'already on PATH (or PATH update skipped)'}")

    step(4, "Claude Code skills")
    for name in SKILLS:
        copytree(os.path.join(SRC_SKILLS, name), os.path.join(CLAUDE, "skills", name))
        say(f"   {name}")

    step(5, "Claude Code hook")
    os.makedirs(os.path.join(CLAUDE, "hooks"), exist_ok=True)
    for n in ("xl-post-write.js", "xl-post-write.py"):
        t = os.path.join(CLAUDE, "hooks", n)
        if os.path.isfile(t):
            os.chmod(t, 0o666)
        shutil.copy2(os.path.join(SRC_HOOKS, n), t)
    say("   xl-post-write.js / .py")

    step(6, "settings.json")
    register_hook()
    say(f"   PostToolUse entry: {hook_command()}")

    step(7, "Checks")
    ok = doctor()
    say("\nDone. Claude Code can use xl right away (just ask it about a workbook).")
    say("To type `xl` yourself, open a NEW terminal from the Start menu: a terminal window that was")
    say("already open, or one opened from an app that was already running, keeps the old PATH.")
    return 0 if ok else 2


def uninstall() -> int:
    say("Removing xl")
    unregister_hook()
    say("   settings.json entry removed")
    for p in (os.path.join(CLAUDE, "hooks", "xl-post-write.js"), os.path.join(CLAUDE, "hooks", "xl-post-write.py"),
              os.path.join(BIN, "xl"), os.path.join(BIN, "xl.cmd")):
        if os.path.isfile(p):
            os.remove(p)
            say(f"   removed {p}")
    for d in [os.path.join(CLAUDE, "skills", n) for n in SKILLS] + [XL_HOME]:
        if not os.path.isdir(d):
            continue
        if is_dev_checkout(d):
            say(f"   kept {d} (git checkout)")
            continue
        shutil.rmtree(d, onexc=lambda fn, p, _e: (os.chmod(p, 0o666), fn(p)))
        say(f"   removed {d}")
    say("Done. Python packages were left in place (pip uninstall openpyxl pywin32 pillow to remove them).")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("action", choices=["install", "uninstall", "doctor"])
    a = ap.parse_args()
    if a.action == "install":
        return install()
    if a.action == "doctor":
        return 0 if doctor() else 2
    return uninstall()


if __name__ == "__main__":
    sys.exit(main())
