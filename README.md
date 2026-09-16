<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/hero-dark.png">
  <img src="docs/readme/hero-light.png" width="100%" alt="xlsx is a Claude add-on that forces AI to work on Excel the way an analyst would. By default Claude in Excel is slow, hardcodes values, forgets to name the reference cell, misses mistakes and burns a lot of tokens. BCN TMT Labs fixes that: it cuts response time by 50 to 75%, forces Claude to never hardcode values by default, cites the reference cell for every number, doesn't miss mistakes, and cuts the token cost of working in Excel by about 10 to 20%.">
</picture>

**Needs:** Windows · desktop Excel · [Claude Code](https://claude.com/claude-code) &nbsp;|&nbsp;
[How to install](#how-to-install) · [Commands](#commands) · [Install guide](INSTALL.md) · [Overview page](https://ignacioelias.github.io/xl/)

## The root of the problem

Claude Code has no built-in way to read a spreadsheet. When you ask about a workbook, it writes a
short Python script using **openpyxl** (a library that unzips the .xlsx and reads its XML), runs
it, reads what it prints, and then the script ends. For the next question it writes a new script,
which starts a new Python process that loads the whole file again. On a 50-sheet business plan
that takes 13 to 18 seconds, every time.

openpyxl can see formulas as text, and it can see the values Excel stored the last time someone
saved the file. It cannot calculate. So when a question needs a number that isn't already stored
("what if churn were 5%?"), plain Claude either rebuilds the model's logic in Python, where it can
easily get the arithmetic wrong, or trusts stored values that may be out of date.

**xlsx changes two things.** The workbook opens once and stays in memory, so each follow-up question
is answered in about a second. And any number that matters is calculated by Excel itself, running
hidden in the background, so the answer is what Excel would show, with the cell reference attached
(for example `Summary!D42`) so you can check it yourself.

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/diagram-dark.png">
  <img src="docs/readme/diagram-light.png" width="100%" alt="Without xlsx, Claude Code writes a script, a new Python process does its own arithmetic and fully reloads the 10 MB workbook (13 to 18 s), then exits and loses its memory; this repeats for every question. With xlsx, the short-lived xl command talks over TCP to the xlsx kernel, which stays up and keeps variables; the kernel reads from an openpyxl copy in memory in about a second, computes in a private hidden Excel, and runs the open test that gives you the Excel answer.">
</picture>

*With xlsx, the kernel loads the file once and reloads it only when the file's date or size
changes. Reading goes to openpyxl, and any number that has to be calculated goes to Excel. Before
a file is handed over, a throwaway Excel with alerts switched on opens it to confirm there is no
repair prompt.*

## What you can ask it to do

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/cards-dark.png">
  <img src="docs/readme/cards-light.png" width="100%" alt="What is in this model? Lists sheets, tables, named ranges and sheet links. Where does this number come from? Walks the formulas back or forward. What if churn were 5% instead of 4%? Excel recalculates and reports the outputs; the original is only saved if you ask. Check this model for errors: full recalculation, every error and stale value, 18 checks and an open test. Build the think-cell tab for this bridge: a feed laid out for think-cell with correct values before any recalculation, checked before handover.">
</picture>

Claude picks xlsx up on its own. You don't need to learn any commands.

## Why xlsx matters

Same questions, same model, without and with xlsx.

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/stats-dark.png">
  <img src="docs/readme/stats-light.png" width="100%" alt="Correct answers 90% to 100%. Median time per answered question 35 s to 19 s. Eight exploration questions run by script with no AI 156 s to 44 s. API spend per 10 questions $1.78 to $1.61.">
</picture>

### What the evidence actually shows

A benchmark of 10 questions on one real 50-sheet client plan (636k formulas), each answer checked
against Excel by script. Claude Sonnet ran headless, once with Python and openpyxl only and once
with xlsx.

| Measure | Without xlsx | With xlsx | What it means |
|---|---|---|---|
| Correct answers | 18 / 20 | 10 / 10 final run<br>35 / 35 all runs | Both misses were the same question: "which cells reference `Summary!R45`?" Answering it means scanning every formula in 50 sheets. Plain Python started a background scan that never reported back, while xlsx answers it from its prebuilt index. The accuracy gain comes from this one type of question, the reverse lookup. |
| Turns per question (median) | 3 to 3.5 | 3 final run<br>5 to 8 earlier | Without xlsx, Claude also needed about 3 turns. The drop to 3 with xlsx came from rewriting the skill that teaches Claude the tool. |
| Median time per question | 35 to 39 s | 19 s | Roughly half. |
| Cost per 10 questions | $1.72 to $1.84 | $1.61 | About 10% lower. |
| Time to answer the 10 questions | 610 to 1,021 s | 260 s | 57 to 75% faster. |

Indicative: one model, one model family, one or two runs per condition, September 2026. The model
is confidential, so these runs can't be reproduced from this repository. The scripted load test
behind the 156 s to 44 s figure (8 questions, no AI) is
[`bench/bench_repl_vs_scripts.py`](bench/bench_repl_vs_scripts.py).

## The 18 checks it performs

They mirror errors found in real financial models.

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/checks-dark.png">
  <img src="docs/readme/checks-light.png" width="100%" alt="The 18 checks: a named range pointing at deleted cells; an undocumented link to another workbook; a circular reference hidden by iterative calculation; merged cells; hidden rows or columns; an error value typed as a constant; a formula pattern that breaks mid-row; a hardcoded number between formulas; a number stored as text; a stale cached error; a SUM that stops one row short; overlapping ranges summed twice; a blank cell read as zero; a chart series pointing at a deleted sheet; an external link with an absolute path; SUM or AVERAGE skipping text; one formula mixing two currencies; a currency figure added to a date. Real example: on one business plan the first run found 14 of 15 named ranges broken, more than 10,000 cached errors and a formula-pattern break on row 334 of every cohort sheet.">
</picture>

Fourteen more rules exist in the code but are switched off by default (`OFF_IDS` in
[`xl/lint.py`](xl/lint.py)). `bench/make_fixture.py --check` fails if an active rule stays silent
on its planted defect or a switched-off rule fires.

## How it works

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/readme/flow-dark.png">
  <img src="docs/readme/flow-light.png" width="100%" alt="You ask Claude Code about a workbook, in English or Spanish. Claude Code recognises the task and calls xlsx. The xlsx kernel holds the workbook in memory and finds, reads and traces in about a second. The Excel engine, a hidden separate instance, recalculates, runs what-ifs, renders ranges and test-opens the file.">
</picture>

After Claude edits a workbook, a quick check runs automatically and the findings go back to Claude
before it reports the work as done. Excel runs in a private hidden instance, so your own Excel
windows are never touched. Originals are never saved over by default: edits and what-ifs happen on
working copies, and Claude only changes the original if you tell it to.

xlsx is a standalone command-line program (the command is `xl`) with its own persistent kernel, so it keeps working if a
given chat tool is unavailable. Every command runs the same way when typed by hand in a plain
terminal, although using it that way takes some Python. Claude Code is one way to drive it: the
`xl-repl` skill teaches Claude when to reach for xlsx and gives it the full API before its first
call, and a hook runs a quick check after every edit. It is an independent open-source project,
not an official Anthropic product.

**Under the hood.** The `xl` command is a thin client that passes Claude's script to a long-lived
Python kernel (one per session, local TCP with a token, 3 h idle timeout) where the workbook is
already loaded and variables persist. The kernel reloads a file only when its date or size
changes. Calculation, what-ifs and rendering go to a hidden Excel that xlsx starts itself with
manual calculation, macros off and alerts off; the open test uses a separate Excel with alerts on.
The skill sets four rules: check for errors after every write, take numbers from Excel, cite
`Sheet!Address` for every figure, and never edit an original in place.

## How to install

**You need:** Windows 10 or 11, desktop Excel and [Claude Code](https://claude.com/claude-code).
The installer adds Python 3.14 if it is missing. About 5 minutes.

**1. Open Claude Code in any folder and paste this prompt:**

```text
Install xl on this machine and confirm it works.
1. Clone https://github.com/Ignacioelias/xl into %USERPROFILE%\tools\xl. If that folder already holds a clone of this repository, run git pull there instead. If it exists and is not a clone, stop and ask me.
2. Run: powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\tools\xl\install.ps1" -NoPause
3. It must end with six PASS lines and "All checks passed." On any FAIL, show me the exact output and stop; do not try to fix it yourself.
4. Run `xl help` (fallback: py -3.14 "%USERPROFILE%\tools\xl\xlcli.py" help) and confirm it prints the command list.
5. Tell me in three lines that xl is installed and what I can ask you now. Do not change anything else on my machine.
```

**2. Claude runs the installer for you.** It installs the Python packages, registers the two
Claude Code skills (xl and think-cell feeds) and the hook, adds the `xl` command, and runs six
checks with Excel. You want to see this:

```text
[7/7] Checks
   PASS  Python packages (openpyxl, pywin32, pillow)
   PASS  Excel reachable through COM
   PASS  xl kernel starts and answers
   PASS  Lint fixture: every rule fires on its planted defect
   PASS  Excel opens the fixture without a repair prompt
   PASS  Think-cell feeds: injected values read by Excel, no repair

   All checks passed.
```

**3. Use it.** Start Claude Code in a folder with a workbook and ask: *"What is in this model, and
does it have any errors?"*

To install by hand, see what the installer changes on your machine, update or uninstall, read
[INSTALL.md](INSTALL.md). If a check fails, please
[open an issue](https://github.com/Ignacioelias/xl/issues) with the full output.

## Commands

Everything also works by hand; `xl help` prints the full API.

| Command | What it does |
|---|---|
| `xl exec` | Run Python in the kernel, with the workbook helpers available (the main entry point) |
| `xl describe`, `sheets`, `find`, `find-rows`, `lookup`, `read` | Explore a workbook |
| `xl precedents`, `dependents`, `trace` | Follow formulas backwards or forwards |
| `xl lint` | The 18 structural checks, no Excel needed |
| `xl calc [--verify]` | Full recalculation in Excel; lists error cells and stale cached values |
| `xl scenario`, `sweep` | What-ifs in Excel, nothing saved |
| `xl render [--diff]` | PNG of a range, optionally compared with a baseline |
| `xl check-open` | Opens the file with alerts on and reports any repair prompt |
| `xl inject`, `calcpr` | Write cached values into formula cells and set calculation options, without Excel |
| `xl status`, `stop` | Manage kernels |

Also included, experimental and not covered by the install checks: `xl pptx-lint` and
`xl pptx-render` for PowerPoint decks.

## Tests

```powershell
py -3.14 -m compileall -q xl
py -3.14 bench/make_fixture.py "$env:LOCALAPPDATA\xl\work\fixture_defects.xlsx" --check
py -3.14 bench/inject_check.py "$env:LOCALAPPDATA\xl\work\inject_check"
py -3.14 bench/witan_cases.py
py -3.14 skills/thinkcell-chart-feeds/scripts/selftest.py
```

The README images are rendered from [`docs/index.html`](docs/index.html) by
`py -3.14 docs/readme/render.py`.

## Limits

- Windows, desktop Excel and Python 3.14 only.
- Memory: a 10.7 MB workbook holds about 9.7 GB in the kernel (formulas, values and the reference index).
- Each `xl` call has a floor of about 1.2 s for Python start-up.
- Lint reads the file without recalculating; `xl calc` is the authority.
- `render` goes through the Windows clipboard. Text on the clipboard is restored afterwards; rich content is not.

## Credits and licence

The design follows the public research log by [Witan Labs](https://github.com/witanlabs/research-log).
xlsx is an independent project and is not affiliated with Anthropic or Witan Labs.

Apache License 2.0 · © 2026 Ignacio Elías. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
