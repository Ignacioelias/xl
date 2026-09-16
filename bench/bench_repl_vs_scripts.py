"""Timing benchmark: the old flow (one fresh Python + openpyxl load per question) against the
persistent xl kernel (load once, answer many). Same questions, same workbook, wall-clock.

Usage: py -3.14 bench_repl_vs_scripts.py <workbook.xlsx> [--runs 3]
"""
import argparse
import json
import statistics
import subprocess
import sys
import time
from pathlib import Path

XL = ["py", "-3.14", str(Path(__file__).resolve().parent.parent / "xlcli.py")]

# Each task: (name, old-flow script body [uses WB path var], kernel code)
TASKS = [
    ("list sheets",
     "print([ws.title for ws in wb.worksheets][:5])",
     "sheets(P)[:5]"),
    ("find 'EBITDA' labels",
     "hits=[f'{ws.title}!{c.coordinate}' for ws in wb.worksheets for c in ws._cells.values() if isinstance(c.value,str) and 'EBITDA' in c.value]; print(hits[:8])",
     "find(P, 'EBITDA', limit=8)"),
    ("read Summary!P6:AA14",
     "ws=wb['Summary']; print([[c.value for c in r] for r in ws['P6:AA14']][:3])",
     "read(P, 'Summary!P6:AA14', values=False)[:3]"),
    ("count formulas per sheet",
     "print({ws.title: sum(1 for c in ws._cells.values() if isinstance(c.value,str) and c.value.startswith('=')) for ws in wb.worksheets[:6]})",
     "{ws.title: sum(1 for c in ws._cells.values() if isinstance(c.value,str) and c.value.startswith('=')) for ws in wb(P).worksheets[:6]}"),
    ("defined names",
     "print([(k, v.attr_text) for k, v in wb.defined_names.items()][:6])",
     "[(k, v.attr_text) for k, v in wb(P).defined_names.items()][:6]"),
    ("precedents of Summary!R45",
     "import re; f=wb['Summary']['R45'].value; print(f, re.findall(r'[A-Z]{1,3}\\d+', f or ''))",
     "precedents(P, 'Summary!R45')"),
    ("header row of Engine",
     "ws=wb['Engine']; print([c.value for c in ws[7]][:12])",
     "read(P, 'Engine!A7:L7', values=False)"),
    ("cross-sheet ref count",
     "print(sum(1 for ws in wb.worksheets for c in ws._cells.values() if isinstance(c.value,str) and c.value.startswith('=') and '!' in c.value))",
     "sum(1 for ws in wb(P).worksheets for c in ws._cells.values() if isinstance(c.value,str) and c.value.startswith('=') and '!' in c.value)"),
]


def run_old(path, body):
    script = f"import openpyxl\nwb = openpyxl.load_workbook({path!r})\n{body}\n"
    t0 = time.perf_counter()
    r = subprocess.run(["py", "-3.14", "-c", script], capture_output=True, text=True)
    dt = time.perf_counter() - t0
    return dt, r.returncode == 0, (r.stdout or r.stderr)[:200]


def run_new(path, code, session):
    full = f"P = {path!r}\n{code}"
    t0 = time.perf_counter()
    r = subprocess.run(XL + ["-s", session, "-t", "600", "exec", "-c", full], capture_output=True, text=True)
    dt = time.perf_counter() - t0
    return dt, r.returncode == 0, (r.stdout or r.stderr)[:200]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--runs", type=int, default=3)
    a = ap.parse_args()
    path = str(Path(a.path).resolve())
    session = "bench"
    results = {"old": {}, "new": {}}
    for run in range(a.runs):
        subprocess.run(XL + ["-s", session, "stop", "--force"], capture_output=True)
        for name, old_body, new_code in TASKS:
            dt, ok, out = run_old(path, old_body)
            results["old"].setdefault(name, []).append((dt, ok))
            print(f"run{run} OLD {name:<28} {dt:6.1f}s ok={ok} {out.strip()[:60]!r}", flush=True)
        for name, old_body, new_code in TASKS:
            dt, ok, out = run_new(path, new_code, session)
            results["new"].setdefault(name, []).append((dt, ok))
            print(f"run{run} NEW {name:<28} {dt:6.1f}s ok={ok} {out.strip()[:60]!r}", flush=True)
    subprocess.run(XL + ["-s", session, "stop"], capture_output=True)

    print("\n| Task | old flow (s) | xl kernel (s) |")
    print("|---|---:|---:|")
    tot_old = tot_new = 0.0
    for name, _, _ in TASKS:
        o = statistics.median(d for d, _ in results["old"][name])
        n = statistics.median(d for d, _ in results["new"][name])
        tot_old += o
        tot_new += n
        print(f"| {name} | {o:.1f} | {n:.1f} |")
    print(f"| **total, {len(TASKS)} questions** | **{tot_old:.1f}** | **{tot_new:.1f}** |")
    ok_old = all(ok for v in results["old"].values() for _, ok in v)
    ok_new = all(ok for v in results["new"].values() for _, ok in v)
    print(f"\nall old ok={ok_old}  all new ok={ok_new}  runs={a.runs} (medians)")
    Path(__file__).with_suffix(".json").write_text(json.dumps(results, indent=1))


if __name__ == "__main__":
    main()
