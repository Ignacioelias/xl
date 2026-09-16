---
name: thinkcell-chart-feeds
description: Build or edit an Excel tab meant to be linked into PowerPoint via think-cell — the layout contract (year row, blank spacer row, data), the waterfall shape, the period-labelling discipline, the build sequence that never recalculates the source model, `xl inject` / `xl calcpr` for giving new formula cells a correct cached value without a recalculation, and `tcfeed`, a declarative pipeline (JSON spec of blocks and waterfalls -> built, injected, verified). For a price x volume revenue bridge use the revenue-bridge skill, which writes think-cell-ready blocks itself. Use whenever the user asks to build a chart feed, a think-cell tab, a bridge or waterfall for a slide, or says "thinkcell tab", "chart feed", "ready for think-cell", or in Spanish "pestaña de thinkcell", "prepara el feed para think-cell", "monta el puente/waterfall para la slide". Composes with xl-repl (reading and verification) and, where installed, xlsx-toolkit (general Excel/COM safety).
---

# Think-cell chart feeds

General method for building an Excel tab that think-cell reads to drive a PowerPoint
chart. Distilled from building these tabs on live financial models where the tab has
to reconcile to the model exactly and the workbook cannot be casually recalculated.
Project-specific row maps and workbook facts belong in that project's own skill, not
here. This file holds only what transfers between engagements.

The tooling is part of `xl` (skill `xl-repl`): `xl read` / `xl find` to read the model,
`xl inject` and `xl calcpr` to finish a build without recalculating, `xl check-open` to
prove the file opens. Read `xl-repl`'s `references/api.md` before the first `xl` call.

## The layout contract

**One blank row sits between a year/label row and its data.** think-cell needs the gap
when the range is selected; without it the first data row gets read as part of the header.

```
row n      block title
row n+1    year row (or segment-label row)
row n+2    ← BLANK. the spacer. never write here
row n+3    first data row
...
row n+k    Check (=0)
row n+k+1  ← blank, then the next block
```

Within a block, a blank row also separates sub-groups. Those are cosmetic. The spacer
right after the year row is not: get that one wrong and the chart silently drops or
mislabels a series.

**Cells that carry no value stay genuinely empty.** Don't write `""` into them. A
formula returning an empty string is not the same as an empty cell to think-cell.

## Always state the period. No bare units.

**Every column that carries a number states its period on the face of the block.** A
year label, or an explicit window; never a bare `£m`, `%`, `#` or a rate on its own.

| Wrong | Right |
|---|---|
| `£m` | `Cumulative FY25–FY30 (£m)` |
| `% of total` | `% of cumulative FY25–FY30 total` |
| `% of spend` | `% of FY30 spend (single year)` |

A cumulative figure and a single-year figure over the same metric are routinely an
order of magnitude apart. A reader, or a chart linked to an ambiguous header, cannot
tell which one they're looking at.

- **Where two period bases sit side by side, label both and say they differ.**
- **Annual and cumulative blocks never share a header row.** Give each its own, even if
  that means repeating the year row a few rows down.
- **A window stated in a header must be provable in a check row**, e.g.
  `SUMPRODUCT((years>=Y0)*(years<=Y1))` against the expected column count, so a gap, a
  duplicate, or a stray text year in the model's own year row gets caught, not shipped.
- If building this with a script that formats a literal `%` into a string, remember a
  Python `%`-format needs the percent sign doubled (`%%`), or use `.format()`/an
  f-string instead. Read the header back after writing: a doubled `%%` that ships is a
  real defect that has shipped before.

## Stacked waterfall layout

**Rows are the stack segments; columns are the waterfall steps.**

```
        B          C        D       E       F        G        H
11   Waterfall   Open    Step 1  Step 2  Step 3      ...     Close     ← headers
12   ← BLANK SPACER, never write here
13   Segment A   ####                            ####
14   Segment B   ####            ####
15   Segment C   ####                    (###)
...
19   Check: opening stack (=0)
20   Check: closing total (=0)
```

- **The chart range covers the whole block.** Column C (or the first data column)
  carries the whole opening stack; each increment sits in its own segment row in its
  own column; the last data column is the closing bar.
- **The bottom row of the closing column carries `e`**, think-cell's computed-total
  marker. It is literal text, so the cell's number format must end `;;@` or the literal
  `e` will not display alongside the numeric formats used elsewhere in the block.
- **Checks go below, outside the chart range**: the opening stack must tie to the
  source total at the opening period, and opening + all increments must tie to the
  source total at the closing period.
- If the opening/closing periods are parameters, make them blue input cells and drive
  the header formulas off them. A waterfall built this way re-cuts for any date pair
  without rebuilding the block.

## The component rule

**Splits follow the model, not the deck.** Where a printed deck shows fewer components
than the source model carries, put every model component on the feed tab, under the
model's own labels. Don't pre-aggregate to match what a slide happened to show. If the
deck shows two kinds of a cost and the model has seven, the tab gets seven, and a
`Check (=0)` row ties the components to the model's own total. A grouping judgement
made silently on the tab is a place a real gap can hide; a check row makes that
impossible.

## Never recalculate the source workbook to build the feed

On a large, fragile financial model, a full recalculation can rewrite values elsewhere
in the workbook that a deliverable already reconciles to: the fix for one block breaks
the tie somewhere else. Treat "don't recalculate" as the default on any workbook you
did not build from scratch, and build the feed so it reads correctly without ever
forcing one.

- **Take a copy of the workbook at the start of the session and keep it.** A workbook
  open in a live Excel session with AutoSave on can recalculate and persist that
  silently. When a build finds values changed, diff the session-start copy against the
  build's own backup before concluding anything: that is what separates damage the build
  caused from damage it inherited, and the copy is the clean base for recovery. Before
  touching the file, confirm no `EXCEL.EXE` with a window holds it.
- **`Application.CalculateBeforeSave` defaults to `True` and fires even in manual
  calculation mode.** Set it `False` before any save and assert it is still `False` at
  the save. It has silently recalculated and broken an already-verified build at the
  save step alone. It is an Application property only: setting it on a Workbook raises
  `AttributeError`, so a script that does so never ran as written.
- **`Application.Calculation` cannot be set until a workbook is open.** Open a seed with
  `Workbooks.Add()` first, then set manual, then open the model with `UpdateLinks=0`.
- A newly written formula may show a stale value or zero until something calculates it.
  Excel sometimes evaluates it at write time even in manual mode, and sometimes does not;
  rely on neither. Give it the correct cached value with `xl inject` (below) and verify, so
  the tab is chartable immediately and the formula stays live for whenever the model is
  deliberately recalculated later.

### The build sequence that works

```
session-start copy -> backup -> hidden DispatchEx -> Workbooks.Add() -> Calculation = manual,
CalculateBeforeSave = False -> open the model (UpdateLinks=0) -> write formulas, formats,
labels; calculate NOTHING (no F9, no Sheet.Calculate, no Range.Calculate) -> save -> close
-> quit -> xl inject the computed values -> xl calcpr (restore what the backup had)
-> verify (below)
```

- Use a dedicated hidden instance (`DispatchEx`), never an attach to the user's session.
  Close without saving and quit in a `finally`.
- Read and verify with bulk `Range.Value2` arrays, never cell by cell.
- Never make a second COM call while a script of your own still holds Excel. Never pipe a
  long COM script through `head`: redirect it to a log and read the log.
- Only headless `EXCEL.EXE` processes (no window handle) can be yours to clean up. One
  with a window is the user's session.

### Injecting cached values: `xl inject`

Compute each block's values independently from the model's cached values (`xl read`,
`xl lookup`, or `wb(P, values=True)` inside `xl exec`), using the same definition as the
formula you wrote (same `SUMIFS` criteria, same window, same `INDEX/MATCH` offset), then:

```bash
xl inject "C:/work/model.xlsm" "Thinkcell" --from values.json      # {"C570": 123.4, "D12": "FY25"}
xl inject "C:/work/model.xlsm" "Thinkcell" --set C570=123.4 --set D571=88.2
xl calcpr "C:/work/model.xlsm"                                     # read calc settings
```

Inside `xl exec`: `inject_cached(P, "Thinkcell", values)` and `calcpr(P, set={...})`.

- Only cells that already hold a formula are patched; other cells are reported as
  `no_formula`, cells not in the file as `missing`. Exit code 2 means one of those lists
  is not empty: read it, don't ignore it.
- In place is allowed on a work copy only; anything else needs `--force` (a backup goes to
  the xl work dir first) or `-o` to write elsewhere. Outputs under OneDrive/Teams sync
  roots are refused unless forced.
- Numbers, text (`FY25`) and booleans are supported. Every patched cell is read back and
  must hold exactly one `<v>` with the injected value, or the command fails.
- **Read the calcPr warning.** `fullCalcOnLoad="1"`, or an automatic `calcMode` with an
  older `calcId`, makes Excel recalculate on open and silently replace every injected
  value. Fix with `xl calcpr F --set calcMode=manual --unset fullCalcOnLoad`.
- **A save through a `Workbooks.Add()` seed can drop `iterate="1"`** and add
  `calcCompleted="0"` and `calcOnSave="0"`. Read the backup's settings with `xl calcpr BACKUP` and put them back
  with `xl calcpr F --set iterate=1 ...`. Never repair it by opening the workbook.

Why the command exists, so nobody hand-rolls a patcher again: each of these produced a
file Excel opened "Repaired" and read-only, with every block on the tab unusable.

- An uncalculated formula returning an empty result is stored as a self-closing `<v/>`.
  A patcher that tests for the literal `<v>` appends a second value element instead.
- `t="..."` must be stripped from the `<c>` opening tag only; a whole-cell regex also
  eats `<f t="shared">` and breaks every cell sharing that formula.
- A verifier that reads values back with a loose `<v>([^<]*)</v>` finds whichever `<v>`
  comes first and passes a broken file. That is what made the bug invisible to its own
  check.

`xl inject` handles all three and is proven end to end by the install check
(`bench/inject_check.py`), including a controlled failure on the old logic.

### Verification after any build (in this order)

1. **Structure and values**: read the built block back (`xl read "Sheet!A1:X40"`): the
   spacer rows are empty, headers carry their periods, every `Check (=0)` row reads 0,
   the injected cells show your computed values, nothing above the new blocks moved.
2. **think-cell names**: `Worksheet.Names` (not `Workbook.Names`) still resolve to the
   intended blocks.
3. **Calc settings**: `xl calcpr F` shows what the model had before the build, with no
   warning.
4. **The file opens**: `xl check-open F`. This is the only check that catches an invalid
   package: the XML checks have passed on a file Excel refused to open. Exit 4 means not
   verified.
5. **Formula and injected value agree**: on a *work copy*, `xl calc COPY --verify` and
   filter `changed` to the feed sheet. It must be empty: a recalculation reproduces the
   injected values exactly. Changes elsewhere in the model are the model's own staleness,
   not the feed's, and are the reason the delivered file is not recalculated. `changed` is
   capped at 200 cells workbook-wide: if `changed_count` reaches 200, the check is
   inconclusive for the feed sheet; say so rather than reporting it as passed.
6. Strip any `[trash]/*.dat` parts an Excel save leaves behind. They are undeclared in
   `[Content_Types].xml` and are not supposed to be in the package.

## The pipeline: declare the blocks, let `tcfeed` build them

For new blocks on an existing model, write a JSON spec and run `scripts/tcfeed_cli.py`
instead of hand-writing a COM script. It applies every rule above: the spacer row, period
labels that state their period, the waterfall shape with the literal `e`, `Check (=0)` rows,
literal years in formulas, format cloning from donor rows, native row insertion that keeps
think-cell's names pointing at the right rows, no recalculation, `xl inject`, and the backup's
`calcPr` put back.

```bash
S="$HOME/.claude/skills/thinkcell-chart-feeds/scripts"
py -3.14 "$S/tcfeed_cli.py" plan spec.json                 # dry run: rows, formulas, computed values. Writes nothing
py -3.14 "$S/tcfeed_cli.py" run  spec.json --engine-check  # backup, compute, build, inject, calcPr, verify
xl -s tcpipe stop
```

- **Spec:** grids (a model sheet's year row and columns), then blocks. `standard` blocks take
  periods and rows built from `sumifs`, `sumifs_label`, `window`, `colcount`, `index_match`,
  `year_end`, `link`, `sum_rows`, `combine`, `tie`, `ratio`. `waterfall` blocks take open/close
  years, segments, steps and the model's total. Anchors: `insert_before` a label, `append`, or
  a `row`. `scripts/examples/sample_spec.json` is a complete, tested example.
- **Always read `plan` first.** Every value the run will inject is computed before anything is
  written, from the model's cached values, with the same definition as the formula. Text that
  Excel would coerce differently by locale stops the run instead of producing a guess.
- **Run on a work copy.** The run refuses a synced original (without `--force`) and a file open
  elsewhere, keeps a backup in its run folder (under xl's work dir), and restores it if the
  build, inject or calcPr step fails.
- **Reading the model:** through the xl kernel below 8 MB, streaming openpyxl above (a full
  kernel load of a very large model costs gigabytes of memory). The selftest proves both
  readers return the same values.
- **Verification is part of the run:** layout, injected values, checks at zero, think-cell
  names, the rest of the feed sheet and every other sheet unchanged, no new errors or `#REF!`,
  no merged cells, `calcPr`, no `[trash]` parts, `xl check-open`, and with `--engine-check` a
  recalculated copy that reproduces every injected value.
- **Not covered:** moving existing blocks (procedure below), repointing a feed workbook to a new
  model vintage, relinking PowerPoint.
- `tcfeed` calls a few internal helpers of `xl` (`xl.excelcom`, `xl.cached`, `xl.paths`,
  `xl.procs`) through `scripts/tcfeed/xlbridge.py`. After an xl refactor, rerun the selftest.

`scripts/selftest.py` proves the pipeline end to end on a synthetic model, including negative
controls for every check (about two minutes, xl session `tcpipe`, output in
`%LOCALAPPDATA%\thinkcell-chart-feeds\selftest`). Spec format, builder semantics, the
verification table and the lessons from earlier feed tooling (row maps between vintages,
repointing a feed, comparing cases, checks that cannot fail, COM traps) are in
[references/pipeline.md](references/pipeline.md).

## Moving blocks on an existing tab

think-cell links point into a tab as **sheet-scoped hidden defined names**.
`Workbook.Names` cannot see them; use `Worksheet.Names` to find and audit them. Because
they're defined names, Excel re-points them automatically on a **native** row move, so
charts don't need relinking:

1. Record each shape's block, row offset, position and size, then set
   `Shape.Placement = 3` (don't move/size with cells) so anchors don't stretch during
   the move.
2. Write a marker into an unused far column at each block's first row, and re-find
   blocks by searching for the marker rather than doing arithmetic on row numbers that
   have already shifted by an earlier move in the same run.
3. Move with native `Rows("a:b").Cut()` then `Rows(dest).Insert(xlDown)`. A native move
   is what makes Excel re-point defined names, intra-sheet formulas, and shape anchors
   together; a copy/paste of values does not.
4. Clear the markers, re-anchor shapes to each block's new start, restore
   `Placement = 1`.

A plain row *insertion* above existing blocks (what the pipeline does) is different: shapes
below should move with their cells without resizing, so set `Placement = 2` for the insert.

## Formatting conventions

- **Structural zeros** (a period the model doesn't populate yet) get a suppressing
  format like `0;-0;;@` or `0.0%;-0.0%;;@` so `INDEX`/`MATCH` returning 0 reads as blank.
  **Do not** suppress zeros on a revenue or cost block: a real zero there is meaningful
  in a stacked chart and hiding it misleads the reader.
- Waterfall segments typically want a format like `#,##0;(#,##0);;@`: parentheses for
  negative steps, matching standard deck convention, with `;;` suppressing zero and the
  trailing `@` preserving the literal `e` total marker.
- Match the base font, input-cell styling (blue font on a light fill, boxed) and tab
  colour to whatever convention the surrounding model already uses. A feed tab that
  looks different from the rest of the workbook reads as untrusted.
- Clone formats from a donor row of the same block type (`PasteSpecial` formats, then
  `CutCopyMode = False`) rather than rebuilding them by hand.

## Don't screenshot the tab via clipboard

Copying an Excel range to an image goes through the Windows clipboard and will clobber
whatever the user has copied elsewhere (a PowerPoint object, in particular). Verify a
built tab by reading cell values back (`xl read`), or render it with `xl render`, which
restores the user's clipboard text and refuses a paste that doesn't match the range.

## See also

- `xl-repl`: reading, tracing and verifying the model (`xl describe`, `xl find`,
  `xl read`, `xl calc --verify`, `xl check-open`), and the full `xl` API.
- `xlsx-toolkit` (where installed): the COM mechanics of editing a workbook safely,
  sync-root delivery, and the openpyxl traps. This skill is only about the
  think-cell-specific layout and recalculation concerns once you're inside a safe edit
  session.
