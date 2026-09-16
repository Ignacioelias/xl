# Installing xl

**You need:** Windows 10 or 11, desktop Excel (Microsoft 365 or 2019+), and
[Claude Code](https://claude.com/claude-code). Claude Code on Windows uses Git for Windows, so
`git` is usually already there. The installer adds Python 3.14 with `winget` if it is missing.
Allow about 5 minutes.

## The easy way: let Claude Code install it

Open Claude Code in any folder and paste:

```text
Install xl on this machine and confirm it works.
1. Clone https://github.com/Ignacioelias/xl into %USERPROFILE%\tools\xl. If that folder already holds a clone of this repository, run git pull there instead. If it exists and is not a clone, stop and ask me.
2. Run: powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\tools\xl\install.ps1" -NoPause
3. It must end with six PASS lines and "All checks passed." On any FAIL, show me the exact output and stop; do not try to fix it yourself.
4. Run `xl help` (fallback: py -3.14 "%USERPROFILE%\tools\xl\xlcli.py" help) and confirm it prints the command list.
5. Tell me in three lines that xl is installed and what I can ask you now. Do not change anything else on my machine.
```

## By hand

In PowerShell:

```powershell
git clone https://github.com/Ignacioelias/xl "$env:USERPROFILE\tools\xl"
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\tools\xl\install.ps1"
```

A good run ends like this:

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

If a check fails, please [open an issue](https://github.com/Ignacioelias/xl/issues) with the
full output.

## What the installer changes

| What | Where |
|---|---|
| Python packages `openpyxl`, `pywin32`, `pillow` | your user site-packages (`pip install --user`) |
| The xl program | `%USERPROFILE%\tools\xl` (your clone is used as is) |
| Two Claude Code skills, `xl-repl` and `thinkcell-chart-feeds` | `%USERPROFILE%\.claude\skills\` |
| A post-edit hook, `xl-post-write` | `%USERPROFILE%\.claude\hooks\` |
| One `PostToolUse` entry that runs the hook (a backup of the file is kept) | `%USERPROFILE%\.claude\settings.json` |
| The `xl` command | `%USERPROFILE%\bin\xl` and `xl.cmd`, plus that folder on your user `PATH` |

xl keeps its sessions, logs and work copies in `%LOCALAPPDATA%\xl`. It makes no network calls.

## Updating

```powershell
git -C "$env:USERPROFILE\tools\xl" pull
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\tools\xl\install.ps1"
```

Running the installer again refreshes the skills and the hook.

## Uninstalling

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\tools\xl\uninstall.ps1"
```

This removes the skills, the hook, its `settings.json` entry and the `xl` command. It keeps your
git clone (delete `%USERPROFILE%\tools\xl` yourself) and the Python packages.

## Troubleshooting

- **`winget` is blocked or missing.** Install Python 3.14 from
  [python.org](https://www.python.org/downloads/windows/) with the py launcher, then run the
  installer again.
- **`pip` fails behind a corporate proxy.** Run the installer from a terminal where
  `pip install` works (for example with `HTTPS_PROXY` set).
- **`xl` is not found after installing.** Open a new terminal from the Start menu. Windows only
  passes the updated `PATH` to new processes.
- **An Excel add-in pops up a dialog.** xl closes add-in dialogs on its own hidden Excel
  instances. If a check still hangs, disable the add-in and try again.
