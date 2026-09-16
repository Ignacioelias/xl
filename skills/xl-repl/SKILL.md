---
name: xl-repl
description: >
  Persistent workbook REPL and verifiers for Excel and PowerPoint work on Windows
  (`xl` on PATH). The way in is `xl exec`: Python against a workbook already open in a
  long-lived kernel, one rich call per question; read references/api.md before the first call.
  Open a workbook once in a long-lived Python kernel and answer many questions
  against it (describe, find, find-rows, lookup, read, precedents, dependents, trace), run
  what-ifs without saving (scenario, sweep), then verify through a dedicated hidden Excel
  instance (calc with --verify: full recalculation, every error cell, stale cached values;
  render with --diff; check-open: the repair-prompt test) and lint (18 rules: broken
  names, plugs, inconsistent formulas, SUM ranges that miss a row, double counting, mixed
  currency, currency added to a date, aggregates over text, broken chart series, cached errors,
  merged cells, external links). PowerPoint: pptx-lint
  (off-slide, overflow, occluded text, tiny fonts, empty placeholders) and pptx-render.
  Package level, no Excel: inject (cached values into formula cells, for think-cell feeds that
  must read right without a recalculation) and calcpr (read/set calculation settings). Use for
  ANY exploration, QA, what-if or verification of an .xlsx/.xlsm/.pptx: "what is in this
  model", "where is EBITDA", "trace this cell", "what feeds / what depends on", "what if X
  were 5%", "does the model have errors", "recalculate and check", "render this range", "is
  the file safe to open", "check the deck for overflow", or in Spanish "qué hay en este
  modelo", "dónde está", "traza esta celda", "qué pasa si", "recalcula y comprueba",
  "renderiza este rango", "revisa la deck". Any WRITE follows the hard
  rules below; this skill is how you read, question and verify without
  re-opening a 10 MB workbook on every tool call.
---

# xl-repl — read once, ask many, verify in the engine

Design based on Witan Labs' public research log (persistent REPL beat discrete tools 92% to
74%; openpyxl for reading, the real engine for verification; check errors after every write;
never compute outside the engine; cite addresses; https://github.com/witanlabs/research-log).
Code: `~/tools/xl` (https://github.com/Ignacioelias/xl).

## The tool: one entry point

The workhorse is **`xl exec`**: Python run inside a long-lived kernel where the workbook is
already open and every helper below is a global. Everything else (`xl describe`, `xl find`, …)
is a one-operation shortcut for the same call. Do not discover the CLI by trial: **read
[references/api.md](references/api.md) once, before the first call** — argument and result
shapes are not guessable from the names. Then open with one rich call:

```bash
xl exec --stdin <<'XL'
P = "C:/path/model.xlsx"
print(describe(P))                          # sheets, blocks, header rows, names, cross-sheet links
print(find(P, "EBITDA", sheet="Summary"))   # Sheet!Addr  text
print(read_tsv(P, "Summary!D40:AA50"))      # dense view, addresses kept
XL
```

- **Batch independent reads into one `exec`.** One rich call beats five small ones; `P` and
  every variable persist, so never re-read what you already pulled.
- **Exception, what-ifs:** two calls on purpose — locate and confirm the output cell first,
  then `scenario()` — so you review what you found before changing anything.
- **Output is capped at 20k chars.** Print selectively: slice tables, `count()` instead of
  listing, filter traces by label. `--json` is uncapped when you truly need everything.
- Use `xl exec -f file.py` when the code has backslashes (Git Bash heredocs can strip
  them); `-c "..."` for one-liners.

## Why it exists

Every Bash call is a fresh Python. Loading a 10 MB model takes 13 s, so ten questions cost two
minutes of loading before any thinking. `xl` keeps a kernel alive across tool calls: the
workbook loads once and each further question costs about a second. The kernel also owns one
hidden Excel instance for recalculation, what-ifs and rendering, so verification uses the real
engine, never Python arithmetic. Measured on a real 50-sheet business plan (636k
formulas): cold describe 18 s, warm 1 s, full recalculation 11 s. Benchmark of 8 exploration
questions, medians of 2 runs (`~/tools/xl/bench`): fresh-Python-per-question 155.9 s, kernel
44.0 s, of which 35 s are the two one-off loads; steady state ~1.4 s per question against ~20 s.

## Hard rules (these are the whole point)

1. **After any write, `xl calc <file>` and read the error list before saying "done".** A file
   with 200 `#VALUE!` cells looks identical to a clean one from openpyxl. `--verify` also lists
   cached values that would change, i.e. the file was saved stale.
2. **Numbers come from the engine.** Write the formula, calc, read the value back. For a
   what-if, `scenario`/`sweep` change inputs in memory and never save.
3. **Cite `Sheet!Address` for every figure you report.** Unverifiable numbers are not findings.
4. **Never edit an original in place.** `work_copy(path)` gives you a
   scratch copy; `save_as()` and `replace()` refuse sync-root paths.
5. **Never hide, unhide or merge cells.** Read hidden rows through the file.
6. **PowerPoint: never quit the app, always work on a copy.** `pptx-lint`/`pptx-render` copy
   the deck, open it without a window, close it, delete the copy. DispatchEx attaches to the
   user's PowerPoint (it runs a single instance), so nothing here ever calls Quit.

## Quick start (every line here was run on a real model or deck before being written)

```bash
xl describe "C:/path/model.xlsx"                     # sheets, ghost ranges, blocks, headers, links, broken names
                                                     # (`xl open` / `xl load` are aliases: there is no separate load step)
xl find "C:/path/model.xlsx" "EBITDA margin"         # regex over cell text -> Sheet!Addr  value
xl find "C:/path/model.xlsx" "^#VALUE!$" --values --sheet Engine --count   # count matches, no output cap
xl find-rows "C:/path/model.xlsx" "EBITDA" --sheet Summary --context 1
xl lookup "C:/path/model.xlsx" Summary "Total EBITDA" 2030      # cell by row label + column label
xl read "C:/path/model.xlsx" "Summary!P6:AA14"       # values as last saved; --formulas for formulas
xl precedents "C:/path/model.xlsx" "Summary!R45"     # what it reads, with cached values
xl dependents "C:/path/model.xlsx" "Engine!R281"     # who reads it, all sheets (index built once, ~20 s)
xl trace "C:/path/model.xlsx" "Summary!R45" --depth 3            # walk to inputs; --outputs walks dependents
xl scenario "C:/path/model.xlsx" --set "Inputs!R927=3000" --out "Summary!R45" --out "Summary!AO45"
xl sweep "C:/path/model.xlsx" "Inputs!R927" "2600,3000,3400" --out "Summary!AO45"
xl lint "C:/path/model.xlsx"                         # 18 rules; --quick is the single-load hook mode
xl calc "C:/path/model.xlsx" --verify                # exit 0 clean, 2 errors, 3 stale cached values
xl render "C:/path/model.xlsx" "Summary!B6:AO46" -o out.png --diff baseline.png
xl check-open "C:/path/model.xlsx"                   # open test: alerts ON in a throwaway Excel; names any blocking dialog
xl pptx-lint "C:/path/deck.pptx" [--slides 3,19]     # off-slide, overflow, occluded text, tiny fonts, empty placeholders/charts
xl pptx-render "C:/path/deck.pptx" --slide 19 -o outdir
xl inject "C:/work/model.xlsm" Thinkcell --from values.json   # cached values into formula cells; exit 2 = cells skipped
xl calcpr "C:/work/model.xlsm" --set calcMode=manual --unset fullCalcOnLoad   # no flags: just read the settings
xl exec -c "t = read(P, 'Engine!A7:L7', values=False); t"        # arbitrary Python with state
xl help | xl status | xl stop [--all] [--force]
```

Paths: forward slashes or Git-Bash `/c/...` both work. Quote paths with spaces. Sheet names
with trailing spaces (`'Inputs & Drivers '`) resolve tolerantly. Global options (`-s`, `-t`,
`--max-chars`, `--json`) are accepted anywhere on the line, before or after the subcommand.
Run `xl help` first, not `xl <cmd> --help` one by one: it prints the whole kernel API in one call.
When a question needs a loop over many cells, write it once in `xl exec -c` (openpyxl objects,
`ws._cells`), rather than paging through `find` output.

## Reading and understanding a workbook (the order that costs fewest calls)

1. **Lay of the land** — `describe(P)`: sheets with counts, header rows, cross-sheet links,
   broken names. Hidden sheets and ghost used ranges are flagged here; do not scan for them.
2. **Find things** — `find(P, rx, sheet=...)` over text, `find(..., values=True)` over cached
   results, `find_rows` for the row around a label, `lookup(P, sheet, row_label, col_label)`
   for a value by labels. Labels repeat (a section header above the numeric row): `lookup`
   skips empty hits, but when you read manually pick the row that holds numbers.
3. **Disambiguate label vs formula cell** — the label sits in columns B–F, the numbers start
   at the first year column (`describe` shows the header row). Cite the value cell, not the label.
4. **Trace** — `precedents` / `dependents` for one hop, `trace(..., depth)` for the chain. Traces
   on 600k-formula models are huge: filter, print the first rows, never dump them.
5. **Counting or scanning many cells** — write the loop once inside `exec` over
   `wb(P, values=True)[sheet]._cells.values()`; never page through `find` output and never
   start a background scan. `count()` covers the common case.
6. **What-if** — locate and confirm the output cell in one call; `scenario()` in the next.

## The kernel API (`xl exec -c "..."` or `xl help`; full shapes in references/api.md)

State persists between calls; `_` is the last result. Set `P = "C:/path/model.xlsx"` once.

| call | what |
|---|---|
| `wb(P)` / `wb(P, values=True)` | cached openpyxl workbook: formulas, or values as last saved |
| `sheets(P)`, `describe(P, blocks=8)` | orientation |
| `find(P, rx, values=False, sheet=None, limit=50)` / `find_rows(P, rx, sheet, context)` | search |
| `lookup(P, "Sheet" or "Sheet!A1:H40", row_label, col_label=None)` | cell by human labels |
| `read(P, "Sheet!A1:H30", values=True)` | `Table` of rows (`.tsv()` for tab-separated text) |
| `read_tsv(P, "Sheet!A1:H30")` / `count(P, rx, values, sheet)` | dense TSV with addresses / number of matches, no cap |
| `precedents(P, ref)`, `dependents(P, ref)`, `trace(P, ref, depth, direction)` | dependency graph |
| `scenario(P, {"Sheet!B3": 0.05}, ["Sheet!B40"])`, `sweep(P, ref, [..], [..])` | what-if in Excel, nothing saved |
| `lint(P, quick=False, as_dict=False)` | rule findings |
| `calc(P, save=False, verify=False)` | dict with `errors`, and `changed` when verify |
| `render(P, ref, out=None, diff=None)` | PNG; clipboard route with a PDF fallback when the clipboard is busy |
| `check_open(P)` | open test; auto-closes add-in dialogs, reports repair prompts |
| `pptx_lint(path, slides=None)`, `pptx_render(path, slide, out_dir)` | PowerPoint via COM on a copy |
| `replace(P, search, repl, in_formulas=False, dry_run=True)` | find-and-replace, work copies only |
| `inject_cached(P, sheet, {"C5": 1.5}, out=None, force=False)`, `calcpr(P, set=None)` | cached values / calc settings at the XML level (think-cell feeds; see skill `thinkcell-chart-feeds`) |
| `work_copy(P)`, `save_as(wb_obj, path)`, `cache_info()`, `drop_cache()` | housekeeping |

`Table` is a list of rows that prints aligned; slice it (`t[:5]`) to keep output small.
Output is capped at 20k chars per call (raise with `--max-chars`); `--json` is never capped.

## Sub-agents

Each sub-agent uses its own kernel: `xl -s <name> ...` or `XL_SESSION=<name>`. The kernel is
single-threaded by design (COM apartment). Kernels exit after 3 h idle; `xl stop --all` ends
them and quits their hidden Excel (never the user's).

## The hook

`~/.claude/hooks/xl-post-write.js` (Node front door, ~0.5 s, exits unless the call named an
.xlsx/.xlsm) hands off to `xl-post-write.py` only for a workbook modified in the last four
minutes: quick lint on the `hook` kernel session, findings injected as context. It cannot run
`calc` for you (a recalculation of a 45 MB model is not a side effect a hook should have).
Rule 1 is still yours to execute.

## Lint rules (all proven on `bench/make_fixture.py`, which plants one defect per rule)

| rule | level | meaning |
|---|---|---|
| `defined-name-broken` | error | named range resolves to `#REF!` |
| `cached-error` / `error-constant` | error | error value cached at last save / typed as a constant |
| `chart-series-broken` | error | series points at a sheet that no longer exists |
| `mixed-currency` | error | one formula combines cells formatted in different currencies (`$` with `€`, `N$` with `USD`) |
| `sum-misses-adjacent` | warn | `SUM(B7:B9)` while B10 holds a number: the classic missed row |
| `double-counting` | warn | overlapping ranges added in one formula |
| `hardcode-in-formula-row` | warn | a non-zero constant with formulas on both sides: a plug |
| `inconsistent-formula` | warn | break after >= 2 consistent copies; cohort staircases and first-period columns exempt |
| `number-as-text` | warn | `'1.234,56'` next to numbers |
| `aggregate-over-text` | warn | SUM/AVERAGE/MIN/MAX/MEDIAN over a range holding text or booleans (skipped silently) |
| `currency-date-mix` | warn | `+`/`-` joining a currency-formatted and a date-formatted cell |
| `merged-cells` | warn | merged ranges break sort, filter and copy |
| `external-link-absolute` | warn | link target is an absolute path that breaks when the file moves |
| `empty-cell-coercion` | info | arithmetic reads an empty cell as 0 |
| `external-link`, `iterative-calc`, `hidden` | info | context you need before editing |

**Switched off by default** (not run, not reported): `mixed-percent`,
`chart-series-empty`, `chart-series-errors`, `object-covers-cells`, `column-clipping`,
`data-validation-breach`, `row-format-inconsistent`, `unsorted-lookup`, `duplicate-lookup-keys`,
`broadcast-surprise`, `row-height-clipping`, `chart-invisible`, `chart-legend-crowded`,
`chart-axis-labels-crowded`. Never tell the user a file was checked for these, and do not
hand-roll them unless asked. The code stays in `lint_rules*.py`; `OFF_IDS` in `xl/lint.py` is
the switch. The fixture keeps their planted defects and `--check` fails if any of them fires.

Lint reads the file through openpyxl and does not recalculate; `calc` is the truth. Quick mode
skips the rules that need the values workbook (cached errors, aggregates over text, empty-cell
coercion).

PowerPoint rules: `off-slide`, `text-overflow` (BoundHeight vs shape, AutoSize off),
`text-occluded` (>= 50% covered by a later opaque shape), `tiny-font` (< 8 pt), `empty-chart`,
`empty-placeholder` (info). Hidden shapes are ignored. Validated on a 29-slide deck.

## Known limits (append here when you hit one)

- Cold load is openpyxl's cost: ~1.2 s per MB of xlsx. The kernel is what makes it a one-off.
- Python start-up is ~1.2 s on a typical corporate laptop, so each `xl` call has that floor even when warm.
- `render` uses CopyPicture + chart export (max 40,000 cells). The clipboard is shared with the
  user's Office session; it retries four times, then falls back to PDF export + PyMuPDF.
- `calc` opens read-only unless `--save`; it refuses to save when Excel opened the file
  read-only. Errors are capped at 200 cells per run; `--verify` diffs formula cells only.
- `dependents` builds a whole-workbook reference index on first use (~20 s on 636k formulas),
  cached with the workbook.
- **Add-in dialogs block fresh hidden instances.** Found 2026-09-07: "Arixcel Explorer"
  (class `#32770`) froze every COM call; `DisplayAlerts` does not suppress it. A guard thread
  closes any `#32770` dialog on OUR Excel PIDs that is not titled "Microsoft Excel" (repair and
  alert prompts keep that title and are reported, never dismissed).
- Clipped numbers (`#####`) are no longer linted; if the user asks about them, `render` the
  range and look.
- Not built: Google Sheets (not needed), an authoring convenience API (openpyxl and COM through
  `exec` cover it).
- **Accuracy A/B, 2026-09-08** (10 engine-verified questions on a real 50-sheet business plan, headless Sonnet,
  programmatic scoring, 2 runs each): `xl` 20/20 correct; plain Python/openpyxl 18/20 — its
  misses were a whole-workbook dependents scan that blew the 120 s tool timeout (the agent
  answered "scan still in progress"). But the baseline needed a median 3 turns per question
  (one script) against 5–14 for `xl`, because the agent did not read the API first: it tried a
  non-existent `xl open`, put `--json` after the subcommand (usage error), paged `find` output,
  and only then wrote the loop in `exec`. Fixed the same day (aliases, flags anywhere,
  `find --count`, `lookup` skipping empty header-row hits, integer labels). The lesson is Witan's:
  a tool the agent has to discover by trial costs more turns than plain Python; the skill must
  make the first call the right one.
- `lookup` matches the first label hit whose target cell holds a value; a section header that
  repeats the label above the numeric row ("Total Revenue" on rows 12 and 14 of a Summary sheet)
  used to win and return `None`.
- **2026-09-09 review fixes.** `calc(save=True)` and `replace(dry_run=False)` write work copies
  only; any other path needs `force=True` and gets a backup in the work dir first. `--verify`
  never counts volatile cells (NOW, TODAY, RAND) as stale; `volatile_skipped` says how many.
  `calc --json` and `check-open` now return the documented exit codes (2 error cells, 3 stale,
  4 not verified). A kernel that is alive but busy is never replaced: the client reports
  "alive but not accepting connections"; wait or `xl stop --force`, which kills only processes
  whose image and creation time match the session file. A request that outlives its
  `--timeout` is skipped by the kernel, not run late. `render` restores clipboard text and
  refuses a paste whose size does not match the range. Array (CSE) formulas are inspected
  everywhere. Reads never insert cells into the cached workbook.
- Excel's `SaveAs` through COM needs a backslash path; `C:/Users/...` fails with "cannot access
  the file 'C:\//Users/...'". `os.path.normpath` before calling it.
- An openpyxl `ExternalLink` written without `r:id`, `sheetNames` and a `sheetDataSet` makes Excel
  refuse the file outright ("Open method of Workbooks class failed", no repair prompt, so
  `check-open` reports a traceback rather than a dialog).
