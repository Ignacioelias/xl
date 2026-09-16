# xl API reference

The complete `xl` surface: the kernel API you call inside `xl exec` (Python, state persists), the
CLI shortcuts, result shapes, and output limits. **Read this before the first `xl exec` on a
workbook** — the function names are guessable, the argument shapes and result shapes are not.
This file is the *reference*; `SKILL.md` is the *playbook* (when to reach for what, the reading /
what-if / verification workflows, the hard rules).

`xl exec` runs Python in a long-lived kernel (one per `-s SESSION`, default `default`). Workbooks
are loaded once with openpyxl and cached by path+mtime; the kernel also owns one hidden Excel
instance for recalculation, what-ifs and rendering. Nothing in the kernel recalculates: cached
values are "as last saved" until you call `calc` or `scenario`.

## Invocation

```bash
# One rich call: set P once, compose several operations, print what you need
xl exec --stdin <<'XL'
P = "C:/path/model.xlsx"
print(describe(P))                                  # orientation: sheets, blocks, headers, names, links
print(find(P, "EBITDA", sheet="Summary"))           # Sheet!Addr  text
print(read_tsv(P, "Summary!D40:AA50"))              # dense view with addresses
XL

xl exec -c "read(P, 'Summary!R45').tsv()"           # one-liner; P persists from the previous call
xl exec -f script.py                                # when the code contains backslashes (Git Bash
                                                    # heredocs can strip them) or is worth keeping
```

- Global options `-s SESSION`, `-t SECONDS`, `--max-chars N`, `--json` may go anywhere on the line.
- The value of the last expression is printed (like a REPL); `print()` output is captured too.
  `_` is the last result. Output is capped at 20,000 characters per call (raise with
  `--max-chars`; `--json` is never capped). Print selectively: slice tables (`t[:20]`), use
  `count()` instead of listing, filter traces.
- Paths: forward slashes or `/c/...`; quote paths with spaces. Sheet names with trailing spaces
  resolve tolerantly. `open`/`load` are aliases of `describe`; there is no separate load step.

## Reading (openpyxl, cached)

| call | returns |
|---|---|
| `wb(P, values=False, reload=False)` | the openpyxl workbook: formulas as text; `values=True` → cached results at last save. Iterate `ws._cells.values()` (populated cells only), never `iter_rows()` on models with ghost used ranges |
| `sheets(P)` | `Table` rows `[index, name, state, used_range]` |
| `describe(P, blocks=8)` | text map: sheets with value/formula counts, header row, cross-sheet and external refs; per-sheet blocks; defined names with `#REF!` flagged |
| `read(P, "Sheet!A1:H30", values=True)` | `Table` of rows (values); `values=False` → formulas. `Table` prints aligned, `.tsv()` gives tab-separated text, slicing works |
| `read_tsv(P, "Sheet!A1:H30", values=True)` | TSV string with a column-letter header and a row-number first column, so every value keeps its address |
| `find(P, regex, values=False, sheet=None, limit=50)` | `Table` rows `[Sheet!Addr, text]`; `values=True` searches cached results (e.g. `"^#VALUE!$"`) |
| `count(P, regex, values=False, sheet=None)` | int, no cap |
| `find_rows(P, regex, sheet=None, context=0, limit=20)` | matching rows with their populated cells; `context=N` adds N rows each side |
| `lookup(P, "Sheet" or "Sheet!A1:H40", row_label, col_label=None)` | `[Sheet!Addr, value, row_label_cell, col_label_cell]`; labels are regexes (a year works as `2030` or `"2030"`; an unbalanced label falls back to a literal match); skips label hits whose target cell is empty |

## Tracing (tokenizer-based reference index, built once per workbook)

| call | returns |
|---|---|
| `precedents(P, "Sheet!C12")` | `Table`: first row `["formula", addr, text]`, then `["ref", Sheet!Addr, cached_value]` per reference (ranges appear as ranges) |
| `dependents(P, "Sheet!C12", limit=200)` | `Table` of `[Sheet!Addr]` whose formulas read the cell, all sheets (~20 s to index 636k formulas, cached on the workbook object; a reload after a write rebuilds it) |
| `trace(P, "Sheet!C12", depth=3, direction="inputs"\|"outputs")` | `Table` of `[level, Sheet!Addr, kind, formula_or_value, cached_value]`; kind is `formula`, `input` or `range` (a range >12 cells is summarised, not expanded) |

Array (CSE) formulas are inspected like plain ones by precedents, dependents, trace and every lint rule.

## Engine (dedicated hidden Excel; never touches the user's Excel; nothing saved unless asked)

| call | returns |
|---|---|
| `calc(P, save=False, verify=False, force=False)` | dict: `errors` (every error cell with formula), `changed` when `verify` (cached values that a full rebuild would alter = file saved stale; volatile NOW/TODAY/RAND cells are counted in `volatile_skipped`, never as stale), `backup` when saved, timings. `save=True` only writes a work copy unless `force=True`, and always backs the file up into the work dir first. CLI exit codes (text and `--json` alike): 0 clean, 2 error cells, 3 stale |
| `scenario(P, {"Sheet!B3": 0.05}, ["Sheet!B40"])` | `Table`: header `[output, base, scenario, delta]`, one row per output. Inputs set in memory on a read-only open; nothing written |
| `sweep(P, "Sheet!B3", [0.03, 0.05, 0.07], ["Sheet!B40"])` | `Table`: header `[input, *outputs]`, one row per input value |
| `render(P, "Sheet!A1:H30", out=None, diff=None)` | dict `{png, bytes, range, method}` (≤40,000 cells); `diff=baseline.png` adds `diff_png`, `changed_pct`. Goes through the shared clipboard: the user's clipboard text is restored afterwards and a paste whose size does not match the range raises instead of exporting foreign content |
| `check_open(P)` | open test: opens with alerts ON in a throwaway Excel; names any blocking dialog; repair prompt ⇒ not verified. CLI exit 0 opened clean, 4 not verified |
| `lint(P, quick=False, as_dict=False)` | 18 rule ids (`xl.lint.RULE_IDS`; 14 more switched off in `OFF_IDS`, see SKILL.md); `as_dict=True` → `{"summary": {"by_rule": {...}}, "findings": [{"level","rule","where","msg"}, ...]}` |

## Package level (no Excel; for blocks that must read right without a recalculation)

| call | returns |
|---|---|
| `inject_cached(P, "Sheet", {"C5": 1.5, "D5": "FY25", "E5": True}, out=None, force=False)` | dict `{file, sheet, part, patched, no_formula, missing, backup, calcpr, warning}`. Writes the cached value of cells that already hold a formula (numbers, text, booleans); every patched cell is read back and must hold exactly one `<v>`. In place only on a work copy unless `force=True` (backup taken); `out=` writes elsewhere and refuses sync roots unless forced. `warning` is set when calcPr would make Excel recalculate on open (`fullCalcOnLoad`, or automatic mode). CLI exit 0 all patched, 2 some cells skipped or missing |
| `calcpr(P, set=None, out=None, force=False)` | dict `{file, calcpr, warning}`; with `set={"calcMode": "manual", "iterate": "1", "fullCalcOnLoad": None}` rewrites `<calcPr>` (`None` removes an attribute, a missing element is created) and adds `before` and `backup` |

The why and the build sequence are in skill `thinkcell-chart-feeds`.

## PowerPoint (COM on a copy; never quits the app)

`pptx_lint(path, slides=None)` → off-slide, text overflow, occluded text, tiny fonts, empty
placeholders/charts. `pptx_render(path, slide=None, out_dir=None)` → slide PNGs.

## Writing (only through openpyxl/COM inside `exec`; the hard rules in SKILL.md apply)

`work_copy(P)` → scratch copy path (never edit originals). `save_as(wb_obj, path)` refuses
sync-root paths unless `force=True`. `replace(P, search, repl, in_formulas=False, dry_run=True, force=False)`
find-and-replace; a write is accepted on a work copy only (any other path needs `force=True`; a
backup is taken first). After any write: `calc(path)` and read `['errors']`.

## CLI shortcuts (one operation each; the kernel call is the same)

`xl describe|open|load F`, `xl sheets F`, `xl find F RX [--values] [--sheet S] [--limit N] [--count]`,
`xl find-rows F RX [--sheet S] [--context N]`, `xl lookup F TABLE ROW [COL]`, `xl read F REF [--formulas]`,
`xl precedents|dependents F REF`, `xl trace F REF [--depth N] [--outputs]`,
`xl scenario F --set REF=V ... --out REF ...`, `xl sweep F REF "v1,v2" --out REF`,
`xl lint F [--quick]`, `xl calc F [--verify] [--save [--force]]`, `xl render F REF -o out.png [--diff base.png]`,
`xl check-open F`, `xl pptx-lint D [--slides 3,19]`, `xl pptx-render D --slide N -o DIR`,
`xl inject F SHEET [--set CELL=V ...] [--from values.json] [-o OUT] [--force]`,
`xl calcpr F [--set ATTR=V ...] [--unset ATTR ...] [-o OUT] [--force]`,
`xl help`, `xl status`, `xl stop [--all] [--force]`.

## Errors you will see

- `FileNotFoundError: <relative path>` — inside `xl exec` the kernel is a detached process whose
  cwd is the tools dir: pass absolute paths in Python code. CLI subcommands (`xl read model.xlsx`,
  `-o out.png`, `--diff base.png`) resolve relative paths against your shell's cwd.
- `PermissionError: ... is not a work copy` — `calc(save=True)` / `replace(dry_run=False)` on an
  original: use `work_copy()` first, or `force=True` when overwriting in place is the intent.
- `kernel ... is alive but not accepting connections` — it is busy in a long COM call; the client
  never replaces a live kernel (that would orphan its hidden Excel). Wait, or `xl stop --force`.
- `KeyError: row label /.../ not found` — `lookup` label regex missed; `find` the label first.
- `usage: xl ...` exit 2 — unknown subcommand (there is no `xl open` beyond the alias) or a
  flag that belongs to another subcommand.
- A `find` that returns exactly `limit` rows is truncated: raise `--limit` or use `count()`.
