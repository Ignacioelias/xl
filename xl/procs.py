"""Process identity on Windows, pure ctypes (no tasklist/wmic). A PID read back from a session
file is only ours if the process is still the one we recorded: same image name and a creation
time that matches. PIDs are recycled quickly on Windows, and `taskkill /F` on a recycled PID
once meant killing the user's visible Excel."""
import ctypes

STILL_ACTIVE = 259
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_EPOCH_DIFF = 11644473600.0  # seconds between 1601-01-01 and 1970-01-01


def pid_alive(pid) -> bool:
    """Never use os.kill(pid, 0) on Windows: any signal other than CTRL events terminates."""
    if not pid:
        return False
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return False
    try:
        code = ctypes.c_ulong()
        if not k32.GetExitCodeProcess(h, ctypes.byref(code)):
            return False
        return code.value == STILL_ACTIVE
    finally:
        k32.CloseHandle(h)


def process_start_time(pid):
    """Creation time as epoch seconds, or None if the process cannot be opened."""
    if not pid:
        return None
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return None
    try:
        ft = (ctypes.c_ulonglong * 4)()
        ok = k32.GetProcessTimes(h, ctypes.byref(ft, 0), ctypes.byref(ft, 8), ctypes.byref(ft, 16), ctypes.byref(ft, 24))
        if not ok:
            return None
        return ft[0] / 1e7 - _EPOCH_DIFF
    finally:
        k32.CloseHandle(h)


def image_name(pid):
    """Base name of the executable (e.g. 'EXCEL.EXE', 'python.exe'), or None."""
    if not pid:
        return None
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return None
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = ctypes.c_ulong(1024)
        if not k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return None
        return buf.value.replace("/", "\\").rsplit("\\", 1)[-1]
    finally:
        k32.CloseHandle(h)


def processes_named(image="EXCEL.EXE"):
    """{pid: creation_time_epoch} of every process with that image name (Toolhelp32 snapshot)."""
    k32 = ctypes.windll.kernel32
    TH32CS_SNAPPROCESS = 0x2

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [("dwSize", ctypes.c_ulong), ("cntUsage", ctypes.c_ulong), ("th32ProcessID", ctypes.c_ulong),
                    ("th32DefaultHeapID", ctypes.c_void_p), ("th32ModuleID", ctypes.c_ulong),
                    ("cntThreads", ctypes.c_ulong), ("th32ParentProcessID", ctypes.c_ulong),
                    ("pcPriClassBase", ctypes.c_long), ("dwFlags", ctypes.c_ulong), ("szExeFile", ctypes.c_wchar * 260)]

    k32.CreateToolhelp32Snapshot.restype = ctypes.c_void_p
    k32.Process32FirstW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    k32.Process32NextW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    k32.CloseHandle.argtypes = [ctypes.c_void_p]
    out = {}
    snap = k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == ctypes.c_void_p(-1).value:
        return out
    try:
        pe = PROCESSENTRY32W()
        pe.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = k32.Process32FirstW(snap, ctypes.byref(pe))
        while ok:
            if pe.szExeFile.lower() == image.lower():
                out[int(pe.th32ProcessID)] = process_start_time(int(pe.th32ProcessID))
            ok = k32.Process32NextW(snap, ctypes.byref(pe))
    finally:
        k32.CloseHandle(snap)
    return out


def is_ours(pid, image, not_before, slack=15.0):
    """True only if `pid` is alive, runs `image`, and was created no earlier than `not_before`
    (minus slack). This is the test before any taskkill."""
    if not pid_alive(pid):
        return False
    name = image_name(pid)
    if name is None or name.lower() != image.lower():
        return False
    t = process_start_time(pid)
    if t is None or not_before is None:
        return False
    return t >= float(not_before) - slack
