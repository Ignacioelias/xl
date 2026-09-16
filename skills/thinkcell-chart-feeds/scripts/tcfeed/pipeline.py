"""tcfeed CLI: plan | run | compute | verify.

  run:  guard -> plan -> backup -> compute (from the backup) -> build (hidden Excel, no
        calculation) -> drop [trash] parts -> xl inject -> calcPr restore -> verify

A failure during build, inject or calcPr restores the backup onto the workbook (the failed
file is kept in the run folder). A verification failure keeps the built file and says so.
Exit codes: 0 verified, 1 verification failed, 2 spec/plan error, 3 compute refused,
4 build/inject/calcPr failure (workbook restored), 5 guard refused.
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import sys
import traceback

from . import __version__
from . import a1
from . import compute as C
from . import harness
from . import layout as L
from . import package as P
from . import reader
from . import spec as S
from . import verify as V
from . import xlbridge
from . import xlsem as X
from .builders import SpecError


class Log:
    def __init__(self, path=None):
        self.fh = open(path, "a", encoding="utf-8") if path else None

    def __call__(self, *parts):
        msg = " ".join(str(p) for p in parts)
        print(msg, flush=True)
        if self.fh:
            self.fh.write(msg + "\n")
            self.fh.flush()

    def attach(self, path):
        self.fh = open(path, "a", encoding="utf-8")


def _sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_plan(spec_path, source=None):
    spec = S.load(spec_path)
    src = source or spec["workbook"]
    if not os.path.exists(src):
        raise SpecError(f"workbook not found: {src}")
    vals, forms = reader.scan_sheet(src, spec["feed_sheet"])
    plan = L.Plan(spec, L.SheetState(vals, set(vals) | set(forms)))
    return spec, plan


def do_compute(plan, source, backend, session, log):
    needs = plan.needs()
    model = reader.fetch(source, needs, backend, session, plan.opts["reader_mb_limit"])
    log(f"[compute] {len(needs)} source ranges read ({model.backend}), {len(model)} populated cells")
    values, errors, problems = C.compute(plan, model)
    return values, errors, problems


def values_table(plan, values, width=7):
    out = []
    for b in plan.blocks:
        out.append(f"  {b.id}:")
        for pr in b.rows:
            r = b.row_of(pr.offset)
            cells = [(c, values.get((r, c))) for c in range(b.first_col, b.last_col + 1) if (r, c) in values]
            if not cells:
                continue
            shown = "  ".join(f"{a1.addr(c, r)}={_short(v)}" for c, v in cells[:width])
            more = f"  (+{len(cells) - width})" if len(cells) > width else ""
            out.append(f"    {(pr.label or '')[:34]:<34} {shown}{more}")
    return "\n".join(out)


def _short(v):
    if X.is_num(v):
        return f"{v:,.4g}" if abs(v) < 1e5 else f"{v:,.0f}"
    return repr(v)


def expected_calcpr(before, policy):
    exp = dict(before)
    if policy == "restore+manual":
        exp["calcMode"] = "manual"
        exp.pop("fullCalcOnLoad", None)
    return exp


def restore_calcpr(target, expected, force, log):
    cached = xlbridge.cached()
    cur = cached.read_calcpr(target)
    if cur == expected:
        log(f"[calcpr] unchanged by the build: {cur}")
        return {"before": cur, "after": cur, "changed": False}
    changes = {k: v for k, v in expected.items() if cur.get(k) != v}
    changes.update({k: None for k in cur if k not in expected})
    tmp = target + ".tcf_calcpr"
    res = cached.calcpr(target, set=changes, out=tmp, force=force)
    os.replace(tmp, target)
    log(f"[calcpr] build left {cur}; restored {res['calcpr']}")
    return {"before": cur, "after": res["calcpr"], "changed": True}


def inject(target, sheet, values, force, log):
    cached = xlbridge.cached()
    tmp = target + ".tcf_inject"
    res = cached.inject(target, sheet, values, out=tmp, force=force)
    if res["no_formula"] or res["missing"]:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise harness.BuildError(f"xl inject skipped cells: no_formula={res['no_formula'][:10]} "
                                 f"missing={res['missing'][:10]}")
    os.replace(tmp, target)
    log(f"[inject] {len(res['patched'])} cached values written into {res['part']}")
    return {k: v for k, v in res.items() if k != "file"}


def guard(target, force):
    if not os.path.exists(target):
        return f"workbook not found: {target}"
    locked, why = P.is_locked(target)
    if locked:
        return f"refusing: {why}. Close it (never kill an Excel with a window) and retry."
    if xlbridge.paths().under_sync_root(target) and not force:
        return ("refusing to build in a OneDrive/Teams-synced folder: work on a local copy and deliver "
                "through the gate, or pass --force")
    return None


def cmd_plan(args, log, dry_run_of_run=False):
    spec, plan = load_plan(args.spec)
    session = args.session or spec["options"]["xl_session"]
    backend = args.reader or spec["options"]["reader"]
    log(plan.summary())
    if plan.errors:
        return 2
    values, errors, problems = do_compute(plan, spec["workbook"], backend, session, log)
    log("\nCOMPUTED (what xl inject would write; read from the workbook as it is now)")
    log(values_table(plan, values))
    for p in problems:
        log(f"  REFUSED {p}")
    for (r, c), e in sorted(errors.items()):
        log(f"  EXCEL ERROR {a1.addr(c, r)} = {e}")
    chk = [x for x in C.check_rows(plan, values) if not (X.is_num(x[3]) and abs(x[3]) <= plan.opts["check_tol"])]
    for bid, addr, lab, v in chk:
        log(f"  CHECK NOT ZERO {bid}!{addr} {v!r}  ({lab[:50]})")
    lock = guard(spec["workbook"], getattr(args, "force", False))
    log("\nSTEPS a run would take:")
    log(f"  0 guard      {'OK' if lock is None else lock}")
    log(f"  1 backup     copy of the workbook into the run folder")
    log(f"  2 compute    {len(values)} formula cells from cached values ({backend} reader)")
    log(f"  3 build      hidden Excel, manual calculation, CalculateBeforeSave off; "
        f"{sum(1 for b in plan.blocks if b.insert_time_row)} native insertion(s), {len(plan.blocks)} block(s)")
    log(f"  4 package    drop [trash] parts left by the save")
    log(f"  5 inject     xl.cached.inject: {len(values) - len(errors)} cells")
    log(f"  6 calcpr     restore {xlbridge.cached().read_calcpr(spec['workbook'])} "
        f"({spec['options']['calcpr']})")
    log(f"  7 verify     layout, injected, checks_zero, names, untouched_feed, untouched_model, errors, "
        f"merged, calcpr, package, check_open (xl -s {session})" + (", engine" if getattr(args, "engine_check", False) else ""))
    if problems or (errors and not spec["options"]["allow_errors"]):
        return 3
    return 0


def cmd_run(args, log):
    if args.dry_run:
        return cmd_plan(args, log, dry_run_of_run=True)
    spec, plan = load_plan(args.spec)
    target = spec["workbook"]
    opts = spec["options"]
    session = args.session or opts["xl_session"]
    backend = args.reader or opts["reader"]
    why = guard(target, args.force)
    if why:
        log(f"[guard] {why}")
        return 5
    if plan.errors:
        log(plan.summary())
        return 2
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = args.run_dir or opts["run_dir"] or os.path.join(
        str(xlbridge.paths().WORK), "tcfeed", f"{stamp}_{os.path.splitext(os.path.basename(target))[0]}")
    os.makedirs(run_dir, exist_ok=True)
    log.attach(os.path.join(run_dir, "run.log"))
    log(f"tcfeed {__version__} run {stamp}\n  spec     {spec['_path']}\n  workbook {target}\n  run dir  {run_dir}")
    for w in plan.warnings:
        log(f"  WARN {w}")
    backup = os.path.join(run_dir, "backup_" + os.path.basename(target))
    shutil.copy2(target, backup)
    backup_sha = _sha(backup)
    log(f"[backup] {backup}")
    calc_before = xlbridge.cached().read_calcpr(backup)
    exp_calc = expected_calcpr(calc_before, opts["calcpr"])
    info = {"spec": spec["_path"], "target": target, "backup": backup, "session": session, "reader": backend,
            "calcpr_backup": calc_before, "calcpr_expected": exp_calc, "stamp": stamp}
    _dump(run_dir, "run.json", info)
    _dump(run_dir, "plan.json", plan.to_json())

    values, errors, problems = do_compute(plan, backup, backend, session, log)
    if problems or (errors and not opts["allow_errors"]):
        for p in problems:
            log(f"[compute] REFUSED {p}")
        for (r, c), e in sorted(errors.items()):
            log(f"[compute] {a1.addr(c, r)} computes to {e}")
        log("[compute] nothing was written; fix the spec or the model and rerun")
        return 3
    inj = C.inject_map(values, errors)
    _dump(run_dir, "inject_values.json", inj)
    _dump(run_dir, "computed.json", {a1.addr(c, r): X.to_json(v) for (r, c), v in sorted(values.items())})
    log(f"[compute] {len(values)} cells computed; values in {os.path.join(run_dir, 'inject_values.json')}")

    report = {"info": info}
    try:
        report["build"] = harness.build(plan, values, log)
        trash = P.trash_parts(target)
        if trash:
            P.drop_parts(target, trash, target)
            log(f"[package] dropped {len(trash)} [trash] part(s)")
        report["trash_dropped"] = trash
        report["inject"] = inject(target, plan.feed_sheet, inj, args.force, log)
        report["calcpr"] = restore_calcpr(target, exp_calc, args.force, log)
    except Exception as e:  # noqa: BLE001 - restore, then report
        log(f"[run] FAILED: {e!r}")
        log(traceback.format_exc())
        report["failure"] = repr(e)
        if _sha(target) != backup_sha:
            failed = os.path.join(run_dir, "failed_" + os.path.basename(target))
            shutil.copyfile(target, failed)
            shutil.copyfile(backup, target)
            log(f"[run] workbook restored from the backup; the failed file is {failed}")
        else:
            log("[run] workbook unchanged (nothing was saved)")
        _dump(run_dir, "report.json", report)
        return 4

    log("[verify] ...")
    res = V.run(plan, backup, target, values, errors, exp_calc, session=session, deep=args.deep,
                engine=args.engine_check, open_check=not args.no_open_check, log=log)
    report["verify"] = res.items
    report["chart_ranges"] = {b.id: f"{plan.feed_sheet}!{b.chart_range()}" for b in plan.blocks}
    _dump(run_dir, "report.json", report)
    log("\nVERIFICATION")
    log(res.text())
    log("\nthink-cell ranges:")
    for k, v in report["chart_ranges"].items():
        log(f"  {k:<16} {v}")
    if res.failed:
        log(f"\nNOT VERIFIED ({len(res.failed)} failing). The built workbook is kept; the backup is {backup}")
        return 1
    log(f"\nVERIFIED. Backup: {backup}")
    return 0


def cmd_compute(args, log):
    spec, plan = load_plan(args.spec)
    if plan.errors:
        log(plan.summary())
        return 2
    session = args.session or spec["options"]["xl_session"]
    values, errors, problems = do_compute(plan, spec["workbook"], args.reader or spec["options"]["reader"],
                                          session, log)
    for p in problems:
        log(f"REFUSED {p}")
    out = C.inject_map(values, errors)
    if args.out:
        _dump(os.path.dirname(os.path.abspath(args.out)), os.path.basename(args.out), out)
        log(f"{len(out)} values -> {args.out}  (xl inject WORKBOOK \"{plan.feed_sheet}\" --from {args.out})")
    else:
        log(values_table(plan, values))
    return 3 if problems else 0


def cmd_verify(args, log):
    with open(os.path.join(args.run_dir, "run.json"), encoding="utf-8") as fh:
        info = json.load(fh)
    spec, plan = load_plan(info["spec"], source=info["backup"])
    if plan.errors:
        log(plan.summary())
        return 2
    session = args.session or info["session"]
    values, errors, problems = do_compute(plan, info["backup"], args.reader or info["reader"], session, log)
    if problems:
        for p in problems:
            log(f"REFUSED {p}")
        return 3
    target = args.target or info["target"]
    res = V.run(plan, info["backup"], target, values, errors, info["calcpr_expected"], session=session,
                deep=args.deep, engine=args.engine_check, open_check=not args.no_open_check, log=log)
    log(f"VERIFY {target}")
    log(res.text())
    return 1 if res.failed else 0


def _dump(folder, name, obj):
    with open(os.path.join(folder, name), "w", encoding="utf-8") as fh:
        json.dump(obj, fh, indent=1, default=str, ensure_ascii=False)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass
    p = argparse.ArgumentParser(prog="tcfeed", description="declarative think-cell feed blocks")
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(s):
        s.add_argument("--session", help="xl kernel session (default: options.xl_session, 'tcpipe')")
        s.add_argument("--reader", choices=("auto", "stream", "xl"), help="cached-value reader")

    s = sub.add_parser("plan", help="dry run: layout, formulas, computed values, the steps a run takes")
    s.add_argument("spec")
    s.add_argument("--engine-check", action="store_true")
    common(s)
    s = sub.add_parser("run", help="build, compute, inject, restore calcPr, verify")
    s.add_argument("spec")
    s.add_argument("--dry-run", action="store_true", help="same as `plan`")
    s.add_argument("--engine-check", action="store_true",
                   help="also recalculate a throwaway copy (xl calc --verify) and compare")
    s.add_argument("--deep", action="store_true", help="compare every other sheet whatever the file size")
    s.add_argument("--no-open-check", action="store_true", help="skip xl check-open (not verified then)")
    s.add_argument("--run-dir")
    s.add_argument("--force", action="store_true", help="allow a workbook under a OneDrive/Teams sync root")
    common(s)
    s = sub.add_parser("compute", help="the values xl inject would write, from the workbook as it is")
    s.add_argument("spec")
    s.add_argument("-o", "--out")
    common(s)
    s = sub.add_parser("verify", help="re-run verification for a run folder")
    s.add_argument("run_dir")
    s.add_argument("--target", help="verify this file instead of the run's workbook")
    s.add_argument("--engine-check", action="store_true")
    s.add_argument("--deep", action="store_true")
    s.add_argument("--no-open-check", action="store_true")
    common(s)
    args = p.parse_args(argv)
    log = Log()
    try:
        return {"plan": cmd_plan, "run": cmd_run, "compute": cmd_compute, "verify": cmd_verify}[args.cmd](args, log)
    except SpecError as e:
        log(f"SPEC ERROR {e}")
        return 2
