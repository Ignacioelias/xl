import os
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent.parent
BASE = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")) / "xl"
SESSIONS = BASE / "sessions"
LOGS = BASE / "logs"
WORK = BASE / "work"
for _d in (SESSIONS, LOGS, WORK):
    _d.mkdir(parents=True, exist_ok=True)

def _sync_roots():
    """Folders a sync client uploads from, where a workbook must never be written in place.

    - every OneDrive / SharePoint mount point Windows records under
      HKCU\\Software\\SyncEngines\\Providers\\OneDrive\\*\\MountPoint;
    - the organisation folder that holds synced SharePoint libraries (%USERPROFILE%\\<Org>\\...),
      because shortcuts added to it do not always get their own mount point;
    - the OneDrive / OneDriveCommercial / OneDriveConsumer environment variables;
    - anything listed in XL_SYNC_ROOTS (semicolon-separated).
    """
    home = Path.home()
    found = []
    try:
        import winreg
        key = r"Software\SyncEngines\Providers\OneDrive"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(k, sub) as s:
                        found.append(Path(winreg.QueryValueEx(s, "MountPoint")[0]))
                except OSError:
                    continue
    except (ImportError, OSError):
        pass
    for mp in list(found):
        if mp.parent != home and mp.parent.parent == home:
            found.append(mp.parent)
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        if os.environ.get(var):
            found.append(Path(os.environ[var]))
    found += [Path(p) for p in os.environ.get("XL_SYNC_ROOTS", "").split(";") if p.strip()]
    out = []
    for p in found:
        if p not in out and p != home:
            out.append(p)
    return out


SYNC_ROOTS = _sync_roots()


def under_sync_root(p) -> bool:
    p = Path(p).resolve()
    for root in SYNC_ROOTS:
        try:
            p.relative_to(root.resolve())
            return True
        except (ValueError, OSError):
            continue
    return False


def is_work_copy(p) -> bool:
    try:
        Path(p).resolve().relative_to(WORK.resolve())
        return True
    except (ValueError, OSError):
        return False


def guard_write(p, force=False):
    """Never edit an original in place (calc --save, replace): only a work copy is written without
    force=True. Anything else (a synced original, which OneDrive uploads within seconds, or a
    local original outside the work dir) needs force=True, and the caller takes a backup first."""
    if is_work_copy(p):
        return
    if under_sync_root(p):
        if not force:
            raise PermissionError(f"{p} is a synced original (OneDrive/Teams). Use work_copy() and "
                                  f"deliver through the gate, or pass force=True to overwrite it in place.")
        return
    if not force:
        raise PermissionError(f"{p} is not a work copy. Use work_copy() first, or pass force=True to "
                              f"write it in place (a backup is taken in {WORK}).")


def backup_to_work(p):
    """Copy the file about to be overwritten into the work dir; return the backup path."""
    import shutil
    import time
    src = Path(p)
    dst = WORK / f"{time.strftime('%Y%m%d_%H%M%S')}_backup_{src.name}"
    shutil.copyfile(src, dst)
    return str(dst)
