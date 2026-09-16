"""Compute every formula cell of a plan from the model's cached values, with the same
definition the formula uses (builders.py), in dependency order (lazily, with cycle
detection). The result is what `xl inject` writes as the cells' cached values.

Model values are read from the workbook as it is BEFORE the build (the run's backup): the
build never calculates, so the model's cached values are the same after it. Links into the
feed sheet use the sheet's current row numbers for the same reason.
"""
from . import a1
from . import xlsem as X


class ComputeError(RuntimeError):
    pass


class _Failed(Exception):
    pass


def compute(plan, model):
    """-> (values {(row, col): value}, errors {(row, col): XlError}, problems [str])."""
    memo, failed, stack = {}, {}, []
    problems = []

    def ev(r, c):
        if (r, c) in memo:
            return memo[(r, c)]
        if (r, c) in failed:
            raise _Failed(failed[(r, c)])
        ent = plan.cells.get((r, c))
        if ent is None:
            return None
        blk, cell = ent
        if cell.kind == "text":
            return cell.text
        if cell.kind == "number":
            return cell.number
        if (r, c) in stack:
            raise X.Unmirrored("circular reference: " + " -> ".join(a1.addr(cc, rr) for rr, cc in stack))
        stack.append((r, c))
        try:
            v = cell.builder.value(plan.cx(blk, r, c, cell, model.get, ev))
        except X.Unmirrored as e:
            failed[(r, c)] = str(e)
            raise
        except _Failed as e:
            failed[(r, c)] = f"depends on a cell that could not be computed ({e})"
            raise
        finally:
            stack.pop()
        if v is None:
            v = 0.0
        elif X.is_num(v) and v == 0:
            v = 0.0                      # Excel has no negative zero
        memo[(r, c)] = v
        return v

    for r, c, blk, cell in plan.formula_cells():
        try:
            ev(r, c)
        except (X.Unmirrored, _Failed):
            problems.append(f"{blk.id}!{a1.addr(c, r)}: {failed.get((r, c))}")
    values = {k: v for k, v in memo.items() if plan.cells[k][1].kind == "formula"}
    errors = {k: v for k, v in values.items() if X.is_err(v)}
    return values, errors, problems


def inject_map(values, errors):
    """{"C12": value} for xl inject: every formula cell except those holding an Excel error."""
    return {a1.addr(c, r): v for (r, c), v in sorted(values.items()) if (r, c) not in errors}


def check_rows(plan, values):
    """[(block, addr, label, value)] for every computed cell on a check row."""
    out = []
    for b in plan.blocks:
        for pr in b.rows:
            if pr.kind != "check":
                continue
            r = b.row_of(pr.offset)
            for c in range(b.first_col, b.last_col + 1):
                if (r, c) in values:
                    out.append((b.id, a1.addr(c, r), pr.label, values[(r, c)]))
    return out
