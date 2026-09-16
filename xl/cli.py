import argparse
import json
import os
import re
import sys

from . import __version__
from .client import exec_code, list_sessions, ping, stop


def winpath(s):
    if s is None:
        return None
    m = re.match(r"^/([A-Za-z])/(.*)$", s)
    if m:
        return f"{m.group(1).upper()}:/{m.group(2)}"
    return s


def abspath(s):
    """Client-side absolute path. The kernel is a detached process whose cwd is the tools dir,
    so a relative `-o out.png` or `model.xlsx` must be resolved HERE, against the caller's cwd."""
    s = winpath(s)
    if s is None:
        return None
    return os.path.abspath(s).replace("\\", "/")


def _emit(resp, as_json=False):
    if as_json:
        print(json.dumps(resp, ensure_ascii=False, indent=None))
        return 0 if resp.get("ok") else 1
    if resp.get("stdout"):
        sys.stdout.write(resp["stdout"])
        if not resp["stdout"].endswith("\n"):
            sys.stdout.write("\n")
    if resp.get("result") is not None:
        print(resp["result"])
    if resp.get("stderr"):
        sys.stderr.write(resp["stderr"])
    if resp.get("error"):
        sys.stderr.write(resp["error"] if resp["error"].endswith("\n") else resp["error"] + "\n")
    return 0 if resp.get("ok") else 1


def _run(args, code, as_json_call=False):
    if as_json_call and args.json:
        code = f"import json as _j; print(_j.dumps({code}, default=str, ensure_ascii=False))"
        resp = exec_code(code, session=args.session, timeout=args.timeout, max_chars=args.max_chars)
        return _emit(resp, as_json=False)
    resp = exec_code(code, session=args.session, timeout=args.timeout, max_chars=args.max_chars)
    return _emit(resp, as_json=args.json and not as_json_call)


def _val(v):
    """CLI value: int or float when it parses as one, text otherwise."""
    try:
        return float(v) if "." in v or "e" in v.lower() else int(v)
    except ValueError:
        return v


def _run_counted(args, call, fmt):
    """Run `call` in the kernel, print it (text or --json) and exit 2 when the result's
    count expression is non-zero, 1 on an exception."""
    show = "import json as _j; print(_j.dumps(_r, default=str, ensure_ascii=False))" if args.json else f"print({fmt}(_r))"
    resp = exec_code(f"_r = {call}\n{show}\nlen(_r.get('no_formula', [])) + len(_r.get('missing', []))",
                     session=args.session, timeout=args.timeout, max_chars=args.max_chars)
    if resp.get("stdout"):
        sys.stdout.write(resp["stdout"] if resp["stdout"].endswith("\n") else resp["stdout"] + "\n")
    if resp.get("error"):
        sys.stderr.write(resp["error"] if resp["error"].endswith("\n") else resp["error"] + "\n")
        return 1
    return 2 if (resp.get("result") or "0") != "0" else 0


_GLOBAL_WITH_VALUE = {"-s", "--session", "-t", "--timeout", "--max-chars"}
_GLOBAL_FLAGS = {"--json"}


def _hoist_globals(argv):
    """Accept global options anywhere (agents write `xl read f.xlsx A1 --json`); move them
    in front of the subcommand so argparse sees them. Values are taken verbatim."""
    front, rest, i = [], [], 0
    while i < len(argv):
        tok = argv[i]
        if tok in _GLOBAL_WITH_VALUE and i + 1 < len(argv):
            front += [tok, argv[i + 1]]
            i += 2
            continue
        if tok in _GLOBAL_FLAGS or any(tok.startswith(g + "=") for g in _GLOBAL_WITH_VALUE if g.startswith("--")):
            front.append(tok)
            i += 1
            continue
        rest.append(tok)
        i += 1
    return front + rest


def main(argv=None):
    argv = _hoist_globals(sys.argv[1:] if argv is None else list(argv))
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    p = argparse.ArgumentParser(prog="xl", description="persistent workbook REPL + verifiers")
    p.add_argument("-s", "--session", default=os.environ.get("XL_SESSION", "default"),
                   help="kernel session name (sub-agents: use their own)")
    p.add_argument("-t", "--timeout", type=float, default=180, help="seconds to wait for the kernel")
    p.add_argument("--max-chars", type=int, default=20000)
    p.add_argument("--json", action="store_true", help="machine-readable output")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("exec", help="run Python in the persistent kernel")
    g = s.add_mutually_exclusive_group(required=True)
    g.add_argument("-c", "--code")
    g.add_argument("-f", "--file")
    g.add_argument("--stdin", action="store_true")

    s = sub.add_parser("describe", aliases=["open", "load"],
                       help="structural map of a workbook (also loads it into the kernel; `open`/`load` are aliases)")
    s.add_argument("path")
    s.add_argument("--blocks", type=int, default=8)

    s = sub.add_parser("sheets")
    s.add_argument("path")

    s = sub.add_parser("find", help="regex search over cell text")
    s.add_argument("path")
    s.add_argument("pattern")
    s.add_argument("--values", action="store_true", help="search cached results instead of formulas")
    s.add_argument("--sheet")
    s.add_argument("--limit", type=int, default=50)
    s.add_argument("--count", action="store_true", help="print only the number of matching cells (no limit)")

    s = sub.add_parser("read", help="read a range: xl read file.xlsx \"Sheet!A1:H30\"")
    s.add_argument("path")
    s.add_argument("ref")
    s.add_argument("--formulas", action="store_true")

    s = sub.add_parser("precedents")
    s.add_argument("path")
    s.add_argument("ref")

    s = sub.add_parser("dependents")
    s.add_argument("path")
    s.add_argument("ref")
    s.add_argument("--limit", type=int, default=200)

    s = sub.add_parser("trace", help="multi-level precedent/dependent walk")
    s.add_argument("path")
    s.add_argument("ref")
    s.add_argument("--depth", type=int, default=3)
    s.add_argument("--outputs", action="store_true", help="walk dependents instead of precedents")

    s = sub.add_parser("lookup", help="cell by labels: xl lookup f.xlsx Summary \"Total EBITDA\" 2030")
    s.add_argument("path")
    s.add_argument("table")
    s.add_argument("row_label")
    s.add_argument("col_label", nargs="?")

    s = sub.add_parser("find-rows")
    s.add_argument("path")
    s.add_argument("pattern")
    s.add_argument("--sheet")
    s.add_argument("--context", type=int, default=0)
    s.add_argument("--limit", type=int, default=20)

    s = sub.add_parser("scenario", help="what-if: xl scenario f.xlsx --set Sheet!B3=0.05 --out Sheet!B40")
    s.add_argument("path")
    s.add_argument("--set", action="append", required=True, help="Sheet!A1=value (repeatable)")
    s.add_argument("--out", action="append", required=True, help="output ref (repeatable)")

    s = sub.add_parser("sweep")
    s.add_argument("path")
    s.add_argument("input_ref")
    s.add_argument("values", help="comma-separated values")
    s.add_argument("--out", action="append", required=True)

    s = sub.add_parser("pptx-lint", help="PowerPoint lint via COM (opens a copy)")
    s.add_argument("path")
    s.add_argument("--slides", help="comma-separated slide numbers")

    s = sub.add_parser("pptx-render")
    s.add_argument("path")
    s.add_argument("--slide", type=int)
    s.add_argument("-o", "--out-dir")

    s = sub.add_parser("lint", help="openpyxl checks (no Excel needed)")
    s.add_argument("path")
    s.add_argument("--quick", action="store_true", help="single load; skips cached-error scan")

    s = sub.add_parser("calc", help="full recalculation in Excel; lists error cells. Exit 0 clean, 2 error cells, 3 stale (with --verify)")
    s.add_argument("path")
    s.add_argument("--save", action="store_true", help="write recalculated values back (work copies only unless --force; backup taken)")
    s.add_argument("--force", action="store_true", help="allow --save on a file that is not a work copy")
    s.add_argument("--verify", action="store_true", help="also list cached values that changed; exit 3 if any (volatile formulas excluded)")

    s = sub.add_parser("render", help="PNG of a range via Excel")
    s.add_argument("path")
    s.add_argument("ref", nargs="?")
    s.add_argument("-o", "--out")
    s.add_argument("--diff", help="baseline PNG to diff against")

    s = sub.add_parser("check-open", help="open test: open with alerts ON in a throwaway Excel. Exit 0 opened clean, 4 blocked or not verified")
    s.add_argument("path")
    s.add_argument("--wait", type=int, default=90)

    s = sub.add_parser("inject", help="cached values into formula cells, no Excel: xl inject f.xlsm Sheet --set C5=1.5 [--from v.json] [-o out]. Exit 0 all patched, 2 some skipped")
    s.add_argument("path")
    s.add_argument("sheet")
    s.add_argument("--set", action="append", default=[], help="CELL=value (repeatable); non-numbers stay text")
    s.add_argument("--from", dest="from_json", help='JSON file {"C5": 1.5, "D5": "FY25", "E5": true}')
    s.add_argument("-o", "--out", help="write here instead of in place")
    s.add_argument("--force", action="store_true", help="allow an in-place write on a file that is not a work copy (backup taken)")

    s = sub.add_parser("calcpr", help="read or set <calcPr>, no Excel: xl calcpr f.xlsm [--set calcMode=manual] [--unset fullCalcOnLoad]")
    s.add_argument("path")
    s.add_argument("--set", action="append", default=[], help="attr=value (repeatable)")
    s.add_argument("--unset", action="append", default=[], help="attribute to remove (repeatable)")
    s.add_argument("-o", "--out", help="write here instead of in place")
    s.add_argument("--force", action="store_true", help="allow an in-place write on a file that is not a work copy (backup taken)")

    sub.add_parser("help", help="print the kernel API")
    sub.add_parser("status", help="list kernel sessions and cached workbooks")
    s = sub.add_parser("stop")
    s.add_argument("--all", action="store_true")
    s.add_argument("--force", action="store_true")
    sub.add_parser("version")

    args = p.parse_args(argv)
    c = args.cmd
    if c in ("open", "load"):
        c = "describe"
    if args.json:
        args.max_chars = max(args.max_chars, 5_000_000)

    if c == "version":
        print(__version__)
        return 0
    if c == "status":
        rows = list_sessions()
        if not rows:
            print("no sessions")
            return 0
        for info in rows:
            line = f"{info['session']}: pid={info['pid']} port={info['port']} alive={info['alive']}"
            if info["alive"]:
                d = ping(info["session"])
                if d.get("ok"):
                    line += f" excel_pid={d.get('excel_pid')} calls={d.get('count')}"
                    for e in d.get("cache") or []:
                        line += f"\n    {'values ' if e['values'] else 'formulas'} {e['path']}  load={e['load_secs']}s hits={e['hits']}"
                else:
                    line += f"  ({d.get('error')})"
            print(line)
        return 0
    if c == "stop":
        names = [i["session"] for i in list_sessions()] if args.all else [args.session]
        for n in names:
            print(stop(n, force=args.force))
        return 0
    if c == "help":
        return _run(args, "help()")
    if c == "exec":
        if args.code is not None:
            code = args.code
        elif args.file:
            with open(winpath(args.file), encoding="utf-8") as fh:
                code = fh.read()
        else:
            code = sys.stdin.read()
        return _run(args, code)

    path = repr(abspath(args.path))
    if c == "describe":
        return _run(args, f"describe({path}, blocks={args.blocks})")
    if c == "sheets":
        return _run(args, f"sheets({path})")
    if c == "find":
        if args.count:
            return _run(args, f"count({path}, {args.pattern!r}, values={args.values}, sheet={args.sheet!r})")
        return _run(args, f"find({path}, {args.pattern!r}, values={args.values}, sheet={args.sheet!r}, limit={args.limit})")
    if c == "read":
        return _run(args, f"read({path}, {args.ref!r}, values={not args.formulas})")
    if c == "precedents":
        return _run(args, f"precedents({path}, {args.ref!r})")
    if c == "dependents":
        return _run(args, f"dependents({path}, {args.ref!r}, limit={args.limit})")
    if c == "trace":
        return _run(args, f"trace({path}, {args.ref!r}, depth={args.depth}, direction={'outputs' if args.outputs else 'inputs'!r})")
    if c == "lookup":
        return _run(args, f"lookup({path}, {args.table!r}, {args.row_label!r}, {args.col_label!r})")
    if c == "find-rows":
        return _run(args, f"find_rows({path}, {args.pattern!r}, sheet={args.sheet!r}, context={args.context}, limit={args.limit})")
    if c == "inject":
        vals = {}
        if args.from_json:
            with open(abspath(args.from_json), encoding="utf-8") as fh:
                vals.update(json.load(fh))
        vals.update({k.strip(): _val(v) for k, v in (x.split("=", 1) for x in args.set)})
        if not vals:
            p.error("inject needs --set CELL=value or --from values.json")
        return _run_counted(args, f"inject_cached({path}, {args.sheet!r}, {vals!r}, out={abspath(args.out)!r}, "
                                  f"force={args.force})", "fmt_inject")
    if c == "calcpr":
        changes = {k.strip(): v for k, v in (x.split("=", 1) for x in args.set)}
        changes.update({k.strip(): None for k in args.unset})
        return _run_counted(args, f"calcpr({path}, set={changes or None!r}, out={abspath(args.out)!r}, "
                                  f"force={args.force})", "fmt_calcpr")
    if c == "scenario":
        inputs = {k: _val(v) for k, v in (x.split("=", 1) for x in args.set)}
        return _run(args, f"scenario({path}, {inputs!r}, {args.out!r})")
    if c == "sweep":
        vals = [float(x) for x in args.values.split(",")]
        return _run(args, f"sweep({path}, {args.input_ref!r}, {vals!r}, {args.out!r})")
    if c == "pptx-lint":
        slides = [int(x) for x in args.slides.split(",")] if args.slides else None
        return _run(args, f"pptx_lint({path}, slides={slides!r}, as_dict={bool(args.json)})", as_json_call=True)
    if c == "pptx-render":
        return _run(args, f"pptx_render({path}, slide={args.slide!r}, out_dir={abspath(args.out_dir)!r})", as_json_call=True)
    if c == "lint":
        return _run(args, f"lint({path}, quick={args.quick}, as_dict={bool(args.json)})", as_json_call=True)
    if c in ("calc", "check-open"):
        # Both paths (text and --json) go through the same exit-code mapping. Before 2026-09-09
        # `calc --json` exited 0 on 40 #REF! cells and check-open always exited 0.
        if c == "calc":
            call = f"calc({path}, save={args.save}, verify={args.verify}, force={args.force})"
            show = "print(fmt_calc(_r))"
            verdict = "(_r['error_count'], _r.get('changed_count', 0))"
        else:
            call = f"check_open({path}, timeout={args.wait})"
            show = "print(fmt_check(_r))"
            verdict = "(0 if (_r.get('opened') and not _r.get('blocked')) else 1, 0)"
        if args.json:
            show = "import json as _j; print(_j.dumps(_r, default=str, ensure_ascii=False))"
        code = f"_r = {call}\n{show}\n{verdict}"
        resp = exec_code(code, session=args.session, timeout=args.timeout, max_chars=args.max_chars)
        if resp.get("stdout"):
            sys.stdout.write(resp["stdout"] if resp["stdout"].endswith("\n") else resp["stdout"] + "\n")
        if resp.get("error"):
            sys.stderr.write(resp["error"] if resp["error"].endswith("\n") else resp["error"] + "\n")
            return 1
        try:
            bad, changed = eval(resp.get("result") or "(0, 0)")  # noqa: S307 - our own tuple repr
        except Exception:
            return 1
        if c == "check-open":
            return 4 if bad else 0  # 4: not verified (repair prompt, blocked, or did not open)
        if bad:
            return 2  # error cells present (same convention as witan calc)
        if args.verify and changed:
            return 3  # cached values were stale
        return 0
    if c == "render":
        call = f"render({path}, {args.ref!r}, {abspath(args.out)!r}, diff={abspath(args.diff)!r})"
        return _run(args, call if args.json else f"fmt_render({call})", as_json_call=True)
    p.error(f"unknown command {c}")
    return 2
