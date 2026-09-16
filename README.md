# xl

**xl makes Claude Code work on Excel models the way an analyst would.** It keeps the workbook
open between questions, lets Excel do every calculation, and cites the cell behind each number.

Windows only. Needs desktop Excel and [Claude Code](https://claude.com/claude-code).
Overview page: https://ignacioelias.github.io/xl/

## Install

Open Claude Code and paste the prompt in [INSTALL.md](INSTALL.md), or in PowerShell:

```powershell
git clone https://github.com/Ignacioelias/xl "$env:USERPROFILE\tools\xl"
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\tools\xl\install.ps1"
```

The installer adds Python 3.14 if needed, installs the program, two Claude Code skills and a
post-edit hook, and runs six checks against your Excel. [INSTALL.md](INSTALL.md) lists every
change it makes, and how to update or uninstall.

## What you can ask Claude once it's installed

- "What is in this model?" Sheets, main tables, named ranges and which sheets feed which.
- "Where does this number come from?" The formulas that feed a cell, or everything that depends on it.
- "What if churn were 5% instead of 4%?" Excel recalculates with the new input; the file is not saved.
- "Check this model for errors." A full recalculation in Excel, every error and stale value, 18 structural checks, and an open test that catches repair prompts.
- "Build the think-cell tab for this bridge." A feed laid out the way think-cell reads it, with correct values before any recalculation.

Claude picks xl up on its own. You don't need to learn any commands.

## How it works

Without xl, Claude Code writes a fresh Python script for every question, and each script
reloads the whole workbook with openpyxl (13 to 18 seconds on a 50-sheet plan). openpyxl can
read formulas and the values Excel last saved, but it cannot calculate, so the agent either
does the arithmetic itself or trusts stored values that may be stale.

xl changes that in two ways:

1. **A kernel keeps the workbook in memory.** The `xl` command is a thin client. It passes
   Claude's script to a long-lived Python process (one per session, local TCP with a token)
   where the workbook is already loaded and variables persist. Follow-up questions take about
   a second. The kernel reloads a file only when its date or size changes.
2. **Excel does the calculating.** Recalculation, what-ifs and rendering run in a private,
   hidden Excel instance that xl starts itself, so your own Excel windows are never touched.
   A separate throwaway Excel with alerts switched on opens files to confirm there is no
   repair prompt.

The `xl-repl` skill tells Claude when to use xl, gives it the full API before its first call,
and sets the rules: check for errors after every write, take numbers from Excel, cite
`Sheet!Address` for every figure, and never edit an original in place. The hook runs a quick
lint after any edit to a workbook and passes the findings back to Claude.

## Commands

xl is a standalone command-line program, so everything also works by hand. `xl help` prints
the full API.

| Command | What it does |
|---|---|
| `xl exec` | Run Python in the kernel, with the workbook helpers available (the main entry point) |
| `xl describe`, `sheets`, `find`, `find-rows`, `lookup`, `read` | Explore a workbook |
| `xl precedents`, `dependents`, `trace` | Follow formulas backwards or forwards |
| `xl lint` | 18 structural checks, no Excel needed |
| `xl calc [--verify]` | Full recalculation in Excel; lists error cells and stale cached values |
| `xl scenario`, `sweep` | What-ifs in Excel, nothing saved |
| `xl render [--diff]` | PNG of a range, optionally compared with a baseline |
| `xl check-open` | Opens the file with alerts on and reports any repair prompt |
| `xl inject`, `calcpr` | Write cached values into formula cells and set calculation options, without Excel |
| `xl status`, `stop` | Manage kernels |

Also included, experimental and not covered by the install checks: `xl pptx-lint` and
`xl pptx-render` for PowerPoint decks.

## The 18 checks

Broken named ranges · undocumented links to other workbooks · circular references hidden by
iterative calculation · merged cells · hidden rows or columns · error values typed as constants
· formula patterns that break mid-row · hardcoded numbers between formulas · numbers stored as
text · stale cached errors · SUMs that stop a row short · overlapping ranges summed twice ·
blank cells read as zero · chart series pointing at deleted sheets · external links with
absolute paths · SUM or AVERAGE skipping text · formulas mixing two currencies · currency
figures added to dates.

Fourteen more rules exist in the code but are switched off by default (`OFF_IDS` in
`xl/lint.py`).

## Evidence

Ten questions on a real 50-sheet business plan (636k formulas), each answer checked against
Excel by script, with Claude Sonnet running headless:

| Measure | Without xl | With xl |
|---|---|---|
| Correct answers | 18 / 20 | 10 / 10 (35 / 35 across all runs) |
| Median time per question | 35 to 39 s | 19 s |
| API cost per 10 questions | $1.72 to $1.84 | $1.61 |
| Turns per question (median) | 3 to 3.5 | 3 |

Both misses without xl were the same question, "which cells reference this cell?", which needs a
scan of every formula; xl answers it from a prebuilt index. A scripted timing test with no AI
(8 questions, `bench/bench_repl_vs_scripts.py`) took 155.9 s with a fresh process per question
and 44.0 s through the kernel. The model itself is confidential, so these runs can't be
reproduced from this repository. Treat them as indicative: one model, one model family, one or
two runs per condition.

## Tests

```powershell
py -3.14 -m compileall -q xl
py -3.14 bench/make_fixture.py "$env:LOCALAPPDATA\xl\work\fixture_defects.xlsx" --check
py -3.14 bench/inject_check.py "$env:LOCALAPPDATA\xl\work\inject_check"
py -3.14 bench/witan_cases.py
py -3.14 skills/thinkcell-chart-feeds/scripts/selftest.py
```

`make_fixture.py --check` builds a workbook with one planted defect per rule and fails if an
active rule stays silent or a switched-off rule fires.

## Limits

- Windows, desktop Excel and Python 3.14 only.
- Memory: a 10.7 MB workbook holds about 9.7 GB in the kernel (formulas, values and the reference index).
- Each `xl` call has a floor of about 1.2 s for Python start-up.
- Lint reads the file without recalculating; `xl calc` is the authority.
- `render` goes through the Windows clipboard. Text on the clipboard is restored afterwards; rich content is not.

## Credits and licence

The design follows the public research log by [Witan Labs](https://github.com/witanlabs/research-log).
xl is an independent project and is not affiliated with Anthropic or Witan Labs.

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
