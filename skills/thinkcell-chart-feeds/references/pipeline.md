# tcfeed: the declarative feed-block pipeline

`scripts/tcfeed` builds think-cell feed blocks into an existing model from a JSON spec. It
applies the rules in SKILL.md (layout contract, period labels, waterfall shape, no
recalculation, `xl inject`, `xl calcpr`, verification order) so a session does not re-derive
them. Tested on a synthetic model by `scripts/selftest.py` (results at the end of this file).

```
spec.json ─► plan ─► backup ─► compute (cached values) ─► build (hidden Excel, no calculation)
          ─► drop [trash] parts ─► xl inject ─► calcPr restore ─► verify
```

Requirements: `py -3.14` with pywin32 and openpyxl, and the `xl` tool (`XL_HOME`, default
`%USERPROFILE%\tools\xl`). tcfeed imports `xl.cached` and `xl.excelcom` and calls the `xl`
CLI for `check-open` and `calc --verify`; it never modifies xl.

## Quick start

```bash
S="$HOME/.claude/skills/thinkcell-chart-feeds/scripts"
py -3.14 "$S/tcfeed_cli.py" plan    spec.json                  # dry run: layout, formulas, values, steps. Writes nothing
py -3.14 "$S/tcfeed_cli.py" run     spec.json --engine-check   # the whole pipeline on the spec's workbook
py -3.14 "$S/tcfeed_cli.py" compute spec.json -o values.json   # only the values (for a manual `xl inject --from`)
py -3.14 "$S/tcfeed_cli.py" verify  RUN_DIR [--target F]       # re-run the checks for a finished run
xl -s tcpipe stop                                              # free the xl session when done
```

Always run `plan` first and read it: every row, its first formula, the computed values, the
think-cell range of each block, and any check that would not read 0.

Run on a **work copy**, not the synced original. `run` refuses a workbook under a
OneDrive/Teams sync root unless `--force`, and refuses one that another process holds open
(an owner `~$` file, or a refused write handle).

Exit codes: 0 verified, 1 verification failed (built file kept), 2 spec or plan error,
3 compute refused (nothing written), 4 build/inject/calcPr failure (workbook restored from the
backup, failed file kept in the run folder), 5 guard refused.

Run folder (default `%LOCALAPPDATA%\xl\work\tcfeed\<stamp>_<name>`, or `--run-dir`):
`backup_<workbook>`, `plan.json`, `inject_values.json`, `computed.json`, `run.json`,
`report.json` (build evidence, inject and calcPr results, every check, think-cell ranges),
`run.log`.

## The spec

```json
{
  "workbook": "model_copy.xlsm",            // relative to the spec file, or absolute
  "feed_sheet": "Feed",
  "label_col": "B", "first_col": "C",       // defaults for every block
  "grids": {                                // a model sheet's period axis
    "quarterly": {"sheet": "Calc Q", "year_row": 2, "first_col": "E", "last_col": "AB"},
    "annual":    {"sheet": "Calc_A", "year_row": 2, "first_col": "E", "last_col": "N"}
  },
  "options": {"xl_session": "tcpipe", "check_tol": 1e-6, "calcpr": "restore"},
  "blocks": [ ... ]
}
```

JSON has no comments; the `//` above are for reading only. `scripts/examples/sample_spec.json`
is a complete, tested spec.

**Options:** `reader` (`auto` | `stream` | `xl`), `xl_session` (default `tcpipe`; never
`default`), `run_dir`, `calcpr` (`restore` puts back exactly what the backup had;
`restore+manual` also forces `calcMode=manual` and drops `fullCalcOnLoad`), `check_tol`,
`deep_limit_mb` (compare every other sheet's cached values below this size; `--deep` forces
it), `allow_errors` (inject everything except cells that compute to an Excel error, and report
them), `reader_mb_limit` (the `auto` reader uses the xl kernel below this size).

**Row numbers:** `anchor.row` is a final row number. Every other row number in the spec
(donor rows, `link` targets on the feed sheet) is the sheet's **current** number; the plan maps
it past any rows it inserts.

### Anchors

| anchor | effect |
|---|---|
| `{"insert_before": "^Existing block 2"}` | native `Rows.Insert` above the one row whose label (column `col`, default the label column) matches the regex. Excel re-points intra-sheet formulas, the sheet-scoped names think-cell uses, and shape anchors. Must match exactly one row. Insert blocks come first in the spec. |
| `{"append": true, "gap": 1}` | below the last used row (values or formulas), after one blank row |
| `{"row": 120}` | at that row; the build refuses if any target cell is not empty |

### Standard block

```json
{
  "id": "revenue", "type": "standard",
  "anchor": {"append": true},
  "title": "Revenue by segment | FY2021-FY2026 (EUR m)",
  "notes": ["SOURCE - ...", "BASIS - ..."],
  "periods": {"from": 2021, "to": 2026, "label": "FY{yy}", "header": "formula",
              "actual_flag": {"grid": "annual", "row": 3, "equals": 1, "suffix": ["A", "F"]}},
  "donor": {"title": 3, "header": 4, "data": 6, "memo": 9, "check": 8},
  "format": "number",
  "rows": [
    {"key": "seg_a", "label": "Segment A", "formula": {"fn": "sumifs", "grid": "quarterly", "row": 5}},
    {"blank": true},
    {"key": "total", "kind": "memo", "label": "Memo - total (EUR m)", "formula": {"fn": "sum_rows", "keys": ["seg_a", "seg_b"]}},
    {"kind": "check", "label": "Check (=0) - segments tie to the model total",
     "formula": {"fn": "tie", "total": {"fn": "index_match", "grid": "annual", "row": 9}, "parts": ["seg_a", "seg_b"]}}
  ]
}
```

Rows laid out: title, notes, header, **spacer** (always inserted), your rows, `gap_after`
blank rows (default 1). The think-cell range is the header row down to the last `data` row;
`memo` and `check` rows sit below it.

- `periods` makes one column per year. `columns` instead lists them explicitly:
  `[{"year": 2030, "label": "FY{yy} (single year)"}, {"year": 2030, "from": 2021, "label": "Cumulative FY{ff}-FY{yy}"}]`.
  Label tokens: `{yyyy}`, `{yy}`, `{y}`, and for window columns `{from}`, `{ff}`. A window
  column passes `$from` and `$year` to builders.
- `header`: `text` writes the label; `formula` writes `="FY"&TEXT(MOD(y,100),"00")`, with an
  A/F suffix from the model's own flag row when `actual_flag` is given.
- Every column label must state its period (a year, `FY25`, `Q4 26`, `Dec-24`, ...). A bare
  `EUR m` or `%` is a plan error, and so is a `%%` anywhere.
- Row kinds: `data` (series), `memo`, `check` (label must start `Check` and contain `(=0)`),
  `input` (typed numbers in blue with a mandatory `source`; nothing else on the tab is typed).
- `only_first: true` writes the formula in the first column only.
- `key` lets other rows reference the row; `"other_block.key"` reaches another block (same
  column letter).
- `donor` rows are cloned (formats only, no clipboard) before writing; kinds `title`, `note`,
  `header`, `data`, `memo`, `check`, `input`, `blank`. Without a donor, inserted rows are
  cleared of inherited formats and get minimal styling (bold title and header, italic notes
  and checks, blue input cells on a light fill).

### Waterfall block

```json
{
  "id": "bridge", "type": "waterfall", "anchor": {"append": true},
  "title": "Revenue bridge | FY2022 to FY2026 (EUR m)", "notes": ["SOURCE - ...", "BASIS - ..."],
  "open": {"year": 2022}, "close": {"year": 2026},
  "period_inputs": true,
  "format": "waterfall1",
  "segments": [{"key": "a", "label": "Segment A", "value": {"fn": "index_match", "grid": "annual", "row": 5}}, ...],
  "steps": [
    {"label": "Segment A growth", "segment": "a"},
    {"label": "Launch (FY2023 level)", "segment": "c", "formula": {"fn": "index_match", "grid": "annual", "row": 7, "year": "$open+1"}},
    {"label": "Growth after launch", "segment": "c"}
  ],
  "total": {"fn": "index_match", "grid": "annual", "row": 9}
}
```

Rows: title, notes, the period inputs row (when `period_inputs`), header (opening label, step
labels, closing label), spacer, one row per segment, blank, two check rows, gap. Columns:
opening stack, one per step, closing bar.

- Opening column: each segment's `value` at `open`.
- A step without `formula` is `(segment at close) - opening cell`. When a segment has several
  steps, the one without a formula is the residual: `(segment at close) - opening - other steps`.
  A step with `formula` is evaluated at `close`; `$open`, `$close`, `$open+1` are available.
- The closing column is empty except the last segment row, which carries think-cell's literal
  `e`. The format must end `;;@` (the plan refuses one that does not).
- Checks: `SUM(opening column) - total at open` and `SUM(opening:last step) - total at close`.
- `period_inputs: true` writes the two years as blue input cells; the header labels and every
  formula reference them, so the bridge re-cuts for any pair of years without a rebuild. The
  injected values are for the years in the spec: after changing an input, the model must be
  calculated deliberately (or the block rebuilt).

### Number formats

Presets (or any format code): `number` `#,##0.0;(#,##0.0);0.0`, `number0`, `number_sz` and
`number0_sz` (structural zeros suppressed, `;;@`), `pct`, `pct_sz`, `waterfall`
`#,##0;(#,##0);;@`, `waterfall1`, `count`, `check` `#,##0.000;(#,##0.000);0.000`, `year`,
`text`. `General` is refused (Excel rejects it on non-English locales). Suppress zeros only
where the model does not populate a period yet, never on a revenue or cost line.

## Builders and their Excel semantics

Each builder writes one formula idiom and computes the same definition in Python from cached
values (`builders.py`, one class per idiom). Years are written as **literals** (or an absolute
reference to a numeric input cell), never as a pointer at a header label, so relabelling a
header cannot blank a block. Common parameters: `scale` (divide by it), `sign` (`-1` negates;
costs stored negative, shown positive), `year` (default `$year`).

| fn | formula | computed as |
|---|---|---|
| `sumifs` | `SUMIFS(data, years, y)`; several `rows` added term by term | numbers in the columns whose year cell equals `y`. A blank year cell never matches (a formula caching 0 does). Text and booleans in the data are skipped; an error in a matched column propagates |
| `sumifs_label` | `SUMIFS(INDEX(block,0,MATCH(y,years,0)), labels, "label")` | every row of `first_row:last_row` whose label matches (case-insensitive, `*` `?` `~`), in the first column of `y`. Rows are found by label, so a row inserted in the model does not break it. Annual grids |
| `window` | `SUMPRODUCT((years>=y0)*(years<=y1)*data)` | column by column; blank is 0; any text in the data makes the whole result `#VALUE!` (as Excel); text in the year row passes `>=` and fails `<=` |
| `colcount` | `SUMPRODUCT((years>=y0)*(years<=y1))-expect` | the check that a stated window holds exactly the columns it claims |
| `index_match` | `INDEX(data,1,MATCH(y,years,0)+offset)` | first column of `y` plus `offset` (3 = Q4 on a quarterly grid). No match `#N/A`; outside the row `#REF!`; a blank cell reads 0 |
| `year_end` | `LOOKUP(2,1/(years=y),data)` | the last column of `y`: use it for a stock. A blank cell reads 0 (tested) |
| `link` | `Sheet!$C$5` (or `refs` per year) | the cell; blank reads 0; text passes through |
| `sum_rows` | `SUM(C12,C13,C15)` | keyed rows in the same column; text and blanks ignored |
| `combine` | `C12-C15`, `C12-(builder)`, `0.5*C12+C13` | left to right, Excel arithmetic |
| `tie` | `total-SUM(parts)` | the component-rule check |
| `ratio` | `num/den` (optional `multiplier`, `blank_if_zero`) | `#DIV/0!` when `den` is 0, or `""` with `blank_if_zero` (not an empty cell to think-cell) |

Where Excel's result depends on the locale (text that looks like a number, a date or a
percentage used in arithmetic or in a `SUMIFS` criteria range), the computation stops with
`REFUSED` rather than guessing, and nothing is written. A cell that computes to an Excel error
stops the run unless `allow_errors`. Excel's "close to zero" adjustment of a final `+`/`-` is
not mirrored, which is why check rows are compared with `check_tol`.

Readers: `stream` (openpyxl read-only; flat memory on a large model) or `xl` (the kernel's
`wb(P, values=True)`; quick when the model is already loaded there, but a full load of a large
model costs gigabytes of kernel memory). Both return date-formatted numbers as Excel serials
and error cells as errors. The selftest asserts they return identical values.

## Verification

| check | what it proves |
|---|---|
| `layout` | spacer, blank and gap rows empty; every label on its planned row; formula cells hold formulas; header labels state their period |
| `injected` | every formula cell caches exactly the computed value |
| `checks_zero` | every `Check (=0)` cell reads 0 within `check_tol`, and none is wrapped in `IFERROR` |
| `names` | the feed sheet's sheet-scoped names (think-cell's `___thinkcell...`, read from `localSheetId` in the package, equivalent to `Worksheet.Names`): same set as the backup, re-pointed exactly as the insertions require, none broken, hidden flag kept. The build also compares `Worksheet.Names` through COM and refuses to save on a mismatch |
| `untouched_feed` | every existing cell of the feed sheet keeps its value (moved rows compared at their new row); every existing formula equals the backup's with the expected re-pointing; no stray cells |
| `untouched_model` | every cached value on every other sheet unchanged (skipped above `deep_limit_mb` unless `--deep`; reported as SKIPPED, never as passed) |
| `errors` | no new error values; `#REF!` in formula text counted separately from error values |
| `merged` | no merged cells in the new rows |
| `calcpr` | `<calcPr>` equals the backup's (plus the configured override); a recalculate-on-open warning is printed |
| `package` | no `[trash]` parts |
| `check_open` | `xl check-open` exit 0: the only check that catches a package Excel refuses |
| `engine` (`--engine-check`) | a copy recalculated by `xl calc --verify` reproduces every injected value; INCONCLUSIVE when xl's 200-cell `changed` cap is reached |

The build itself asserts, before saving: manual calculation and `CalculateBeforeSave` off at
every stage, anchor row found and moved by exactly the inserted count, every written cell
read back (formula landed, text not coerced, number exact), think-cell names and shape count
unchanged. It records what Excel cached for each new formula before injection and warns when
conditional-format rules multiplied.

## What it does not do

- Recalculate the target, ever. `--engine-check` recalculates a throwaway copy.
- Move existing blocks (use the marker procedure in SKILL.md), relink PowerPoint, or
  repoint a feed workbook to a new model vintage (see the lessons below and skill
  `model-vintage-migration`).
- Date-valued year rows (a year-end date per column): add a numeric year row to the model or
  a mirror first.
- Whole-row or whole-column references in existing formulas are not re-pointed by the
  `untouched_feed` expectation; a mismatch there shows as a failure to read, not a pass.

## Lessons from three generations of feed tooling

The tooling this pipeline replaces went through three builds: blocks written inside a live
model, a standalone feed workbook of mirror tabs rebuilt for a new model vintage, and a later
repoint plus a case comparison. What follows is what those builds paid for and SKILL.md does
not already say.

### Building blocks

- **Write labels and formulas together, from one plan, every time.** One toolset wrote labels
  once and later runs rewrote formulas only. Reordering the plan then put every formula below
  the edit on its neighbour's label; the check row moved with the block and kept reading 0, and
  the only visible symptom was a percentage of 1e18 where a divisor landed on a check row.
  If a partial rewrite is unavoidable, assert plan against sheet (label, tag, formula, blank)
  before and after. `tcfeed` always writes both and `layout` checks both.
- **The script in a tools folder is not necessarily the script that ran.** Compare it with
  the workbook before re-running it.
- **Text written through `Range.Formula` is parsed.** `1.0` becomes 1, `6.10` becomes 6.1, a
  label starting with `-` or `=` becomes a formula. Prefix an apostrophe (tcfeed does for any
  label with a digit, `%`, a leading `= + - @ # (` or TRUE/FALSE) and read it back.
- **`Range.Formula` without the leading `=` stores text**, and silently eats a leading
  apostrophe of the reference. Read one written cell back and assert its type.
- **COM colours are BGR.** `Font.Color = 255` is red; one toolset coded its "blue input" font
  as 255. Convert from RGB explicitly.
- **`CalculateBeforeSave` is an Application property.** Under pywin32, setting it on a
  Workbook raises `AttributeError` (tested); scripts that did so never ran as written.
- **Never set `Application.Calculation` back to automatic while a workbook is open**, for
  instance in a `finally` that runs before the close: that recalculates it.
- **Kill only your own Excel, by PID.** An old cleanup killed every headless `EXCEL.EXE`; on a
  machine running xl kernels or parallel agents, those are other sessions' instances (three
  were running during this pipeline's own test).
- **An Excel instance outlives `Quit` while any COM object is still referenced**, typically a
  Range held by the frames of an exception traceback. Clear the traceback frames and collect
  before waiting on the PID; kill it only if it has no visible window.
- **Re-pasting formats over rows that already carry them duplicates conditional-format rules**
  and can hang the save. Count `FormatConditions` before and after (the build warns).
- **Pasting column widths also pastes hidden state.** Record and restore visibility. openpyxl's
  `column_dimensions` under-reports grouped `<col>` ranges; read widths from the sheet XML.
- **Do not drive `ActiveWindow` in an invisible instance**, and do not minimise a visible one
  mid-build: both crashed Excel during a save. Freeze panes need a visible window.
- **Cloning an `<xf>` in styles.xml** needs `<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>`; a lazy
  `.*?(?:/>|</xf>)` stops at a child's `/>` and corrupts the package.
- **Leave a segment that does not exist yet genuinely empty**, not zero, so think-cell omits it
  instead of drawing a zero-height bar.
- **A typed number needs a register.** Keep a list of every numeric constant on a feed tab that
  is neither a year nor a model link (cell, label, value, block, source) and recount it after
  each build; `input` rows in tcfeed require a `source` for the same reason.
- **Tag the rows a chart reads.** Deriving the tags from the `___thinkcell` names (column A,
  one colour) shows a reviewer which rows move a chart; tcfeed prints each block's range.
- **Excel evaluated new formulas at entry in manual mode** in this pipeline's test (every one
  of 130 cells already cached the computed value before injection). The earlier builds saw
  zeros and stale values, which is consistent with a formula being evaluated once against
  precedents that were themselves stale or written later, and never again in manual mode.
  Do not rely on either: inject and verify.
- **A save through the `Workbooks.Add()` seed rewrote `<calcPr>`** in the test: `iterate="1"`
  dropped, `calcCompleted="0"` and `calcOnSave="0"` added. Restoring the backup's attributes
  exactly removes all three.

### Checks that can and cannot fail

- **A check built from the same expression as its rows proves nothing.** A year-end headcount
  summed over four quarters was four times too large while its check read exactly 0 in every
  year. Reconcile at least one year line by line against something the build did not produce:
  the previous workbook's cached values, a printed figure. Expect the reference to hold a
  different metric or unit on some lines, and say which.
- **A control that cannot fail is worse than none.** Check rows that a rebuild turned into the
  constant 0, or `"OK"` typed as text, report success whatever the data does. Checks must hold
  formulas (`layout`, `checks_zero`).
- **Never wrap a check in `IFERROR`.** `IFERROR((N(#REF!)*...),"")` renders blank, leaves the
  error count unchanged and validates nothing. Count `#REF!` in formula text as its own check.
- **A check that starts one row too low** (omitting the first component) fails only once
  recalculated; read every check after an engine recalculation of a copy (`--engine-check`).
- **Two series that cover different years make a check fail wherever only one side has data.**
  Blank the check outside the common window rather than hiding it.
- **`Value2` returns error values as integers** (`-2146826281` is `#DIV/0!`); a float test misses
  them. A `startswith("#")` test counts unit labels such as `#` as errors.

### Row maps between model vintages

- **Derive by label, prove by value, with a negative control.** Align each sheet's label
  columns between vintages order-preservingly (`difflib.SequenceMatcher(autojunk=False)` over
  normalised labels). Then, for every referenced row, compare the old row with the claimed new
  row over a window where both vintages must agree (the actual years), and rerun the same test
  with the map offset by ±1 and ±4: most rows must flip to suspect, or the test does not
  discriminate.
- **Shifts are piecewise.** One sheet moved +4, then +8, then +12 in successive bands; another
  0, +2, +3. Record bands and never extend a shift beyond the band that proved it.
- **Repeated labels prove nothing** (`Core`, `Total`, `#` repeat dozens of times). Use label plus
  the populated-column fingerprint, or values. Duplicate lines in the model give non-unique
  value matches; accept a row only if the claimed shift is among the shifts that match it and
  equals its band's shift.
- **Rows with no usable actuals are inconclusive, not confirmed.** Say how many, and which are
  carried by neighbours that pin the same shift.
- **Emit the map as data from the code that consumes it.** A map hand-copied next to its writer
  drifts, and a verification written with the same reader as the build shares its bugs.
- **A shared-formula follower carries no formula text** (`<f t="shared" si="39"/>`). A reader
  that matches `<f>(.*?)</f>` sees a constant and writes the old vintage's cached value as a
  hardcode: thousands of frozen cells in one rebuild, visible only as lines that failed to move.
  Reconstruct followers from the master by translating relative references, and count formulas
  that became constants between builds.
- **A self-closing `<row r="9"/>` swallowed the next row's cells** in a regex reader, shifting
  labels by a row; the Δ=0 reconciliation agreed because it used the same reader. Excel
  evaluating a live reference is the independent oracle, not a second parse.
- **Derive a mirror's scope from every tab that reads the mirrors**, not only the feed tab; a
  comparison tab that read two unmirrored rows showed plausible negative variances.
- **Provenance notes that cite model row numbers go stale when rows shift.** Sweep them with the
  same map as the formulas.

### Repointing a feed to a new vintage

- **Choose the architecture for how the file travels.** A feed workbook beside the model: mirror
  tabs (model tab names and row numbers, cell-by-cell links, aggregation local). A small file
  that travels: direct links with `SUMPRODUCT`/`INDEX` only. `SUMIFS` against a closed workbook
  returns `#VALUE!`; prove portability by opening a copy with every source closed, forcing a
  recalculation and counting errors.
- **On a mirror, link numbers only.** Write text literally and leave blanks blank: a link to an
  empty cell reads 0, which filled label columns with zeros and turned an empty flag row into a
  criterion that matched every quarter.
- **Apply a row shift with a native `Rows.Insert` on the mirror tab**, so Excel re-points every
  reference, including shared-formula followers a text rewriter cannot see. Rewrite `<f>` text
  offline only on sheets proven to hold no shared formulas.
- **The external-link cache is the complete gate.** `xl/externalLinks/externalLinkN.xml` holds
  every source cell the workbook reads, with the value cached at the last refresh. Reconciling
  it cell by cell against the new vintage is total, not a sample; rerun it against the old
  vintage to show it can fail. When the map is the identity, a repoint is a `Target` edit in
  the link's rels part and nothing else, and the file reconciles before any recalculation.
- **Excel rewrites link targets relative to where the workbook was saved.** A build in a scratch
  folder leaves scratch paths; a workbook whose own path is a SharePoint URL gets
  `../../..` targets that resolve nowhere. Force every `Relationship` in
  `xl/externalLinks/_rels/externalLinkN.xml.rels` to the bare URL-encoded filename with
  `TargetMode="External"`, do it last, and keep the feed in the same folder as its sources.
  Opening the build where the sources do not exist makes Excel re-point or duplicate links.
- **Write links in the open-workbook form** (`='[Book.xlsm]Sheet'!$A$1` with the source open);
  the `[1]` index form opens a file dialog and blocks automation.
- **After adding formulas at the XML level, drop `calcChain.xml`** (plus its relationship and
  content-type override) and add the referenced cells to the external-link cache.
- **Static labels on mirror tabs do not refresh.** Sweep them against the new vintage.
- **A shared string is shared.** Editing one `<si>` changes every cell that uses it; list the
  uses first.
- **think-cell finds its data by workbook filename.** Renaming or copying a feed workbook means
  re-pointing every chart by hand. A copy carries the same link identifiers as its original, so
  never keep both open while refreshing the deck.
- **Deliver with `copyfile`, not `copy2`.** A delivery that lands with an older timestamp than
  the file it replaces is invisible to every staleness check. Verify content (cell values), not
  bytes, because Excel and the sync client re-serialise the package, and watch the destination
  for a minute for a sync pull-back.
- **Read a workbook someone has open** with a share-mode snapshot (`CreateFileW` with
  `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`), never by attaching to their session.
- **Every Excel save stamps the file** with the saving user's name (`docProps/core.xml`) and
  the folder it was saved in (`x15ac:absPath` in `xl/workbook.xml`); both were present in every
  workbook this pipeline's test saved. A delivered feed carries them. Check them before a file
  leaves the team.

### Comparing two cases of the same model

- **Materialise a case on a scratch copy**: write the scenario selector exactly, rebuild, and
  assert both the selector and the sub-case it drives read as expected. A sub-case is often
  mapped from the master scenario elsewhere in the model; check the mapping, not the label.
  Never switch or recalculate a file something links to.
- **Same vintage, same addresses; different vintage, never.** A same-address comparison is safe
  only between cases of one vintage. To get an old vintage's figure, compute it from the old
  model: comparing two builds that both link to the new model can only return zero.
- **One sheet can carry several column grids** (a quarterly axis, a block with four columns per
  year, an annual summary with a shorter horizon). Reading a row against the wrong grid returns
  plausible numbers. Aggregate on the grid the model's own formulas use, and use a
  shorter-horizon summary only as a check row, blank beyond its horizon.
- **Stock or flow.** Read a stock at the last column of the year; summing it over four quarters
  returns four times the stock.
- **Prove "identical" per line**: the largest absolute annual difference between the cases, to
  the unit. Tie totals to components on both grids, and separate rate from volume (when the
  numerator is identical, the denominator explains the move).
- **When two related volume lines move in opposite directions, decompose before explaining.**
- **Carry the window the client quotes** (a cumulative column) and full-plan totals, so a line
  that fires outside the window is not reported as nil.
- **Do not force a bridge the model does not close** (a cohort that steps in mid-period);
  report the lines as read.

### Checking a deck against a new vintage

- **Deck numbers are a control, not a target.** Put them on the feed tab as `EXTERNAL CONTROL`
  rows with a `Difference` row beneath; a difference is a model change until shown otherwise.
- **think-cell's plotted values can be read back as label text** on shapes tagged
  `THINKCELLSHAPEDONOTDELETE`. Separate those labels from the prose, which is where most of the
  claims are restated. Slide part names (`slideN.xml`) are not presentation order; take the
  order from `presentation.xml`.
- **Match the window before calling a difference**, state a rounding threshold, and flag a claim
  that cannot be traced to a cell instead of reverse-engineering a number that fits.
- **Copy the deck first and record its timestamp**: a deck under active editing makes a finding
  against a stale copy worse than none.

## Selftest

`py -3.14 scripts/selftest.py` builds a synthetic model in a hidden Excel (quarterly and annual
sheets, a sheet name with a space, an actual/forecast flag, a feed tab with two blocks, two
hidden sheet-scoped names, a shape, a formula pointing below the insertion point; saved with
manual calculation and `iterate="1"`), runs the sample spec with `--engine-check`, asserts every
check and a set of specific facts, then tampers with copies to prove each check can fail
(wrong cached value, check not zero, `fullCalcOnLoad`, text in a spacer, labels out of step, a
broken name, an edited existing cell, an edited model value, a `[trash]` part, a double `<v>`
that Excel refuses), and exercises the CLI paths (verify, compute, the lock guard, three spec
errors, a build failure before the save, a failure after it with restore). It writes
`_test/selftest_result.json` and stops its xl session. About two minutes.
