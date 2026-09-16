"""Client side of the kernel protocol: find or start the session kernel, send one request.

Lifecycle rules (paid for on 2026-09-09):
  - a connection failure is NOT proof the kernel is dead. A kernel busy in a long calc refuses
    connections once its backlog fills; spawning a replacement then orphans it and its hidden
    Excel. Only a dead PID (or a PID that is no longer our python) allows a respawn.
  - spawning takes a lock file, so an agent call and the post-write hook racing on first use
    start one kernel, not two.
  - `stop --force` kills nothing it cannot prove is ours: same image name and a creation time
    that matches the session file. A recycled PID that is now the user's Excel is left alone.
"""
import json
import os
import socket
import subprocess
import sys
import time

from .paths import LOGS, SESSIONS, TOOLS_DIR
from .procs import is_ours, pid_alive, process_start_time
from .server import recv_msg, send_msg

KERNEL_IMAGES = ("python.exe", "pythonw.exe", "python3.exe")
SPAWN_LOCK_STALE = 60.0


def session_info(session):
    f = SESSIONS / f"{session}.json"
    if not f.exists():
        return None
    try:
        return json.loads(f.read_text())
    except Exception:
        return None


def _kernel_alive(info) -> bool:
    """The PID in the session file is alive AND is still the python we started (not recycled)."""
    if not info:
        return False
    pid = info.get("pid")
    started = info.get("started")
    if started is None:
        return pid_alive(pid)
    return any(is_ours(pid, img, started) for img in KERNEL_IMAGES)


def list_sessions():
    out = []
    for f in sorted(SESSIONS.glob("*.json")):
        try:
            info = json.loads(f.read_text())
        except Exception:
            continue
        info["alive"] = _kernel_alive(info)
        out.append(info)
    return out


def _connect(info, timeout):
    s = socket.create_connection(("127.0.0.1", info["port"]), timeout=5)
    s.settimeout(timeout)
    return s


def _lock_path(session):
    return SESSIONS / f"{session}.spawn.lock"


def _acquire_spawn_lock(session):
    lock = _lock_path(session)
    try:
        if lock.exists() and time.time() - lock.stat().st_mtime > SPAWN_LOCK_STALE:
            lock.unlink()
    except OSError:
        pass
    try:
        fd = os.open(str(lock), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        return False


def _release_spawn_lock(session):
    try:
        _lock_path(session).unlink()
    except OSError:
        pass


def _wait_for_kernel(session, wait):
    t0 = time.time()
    while time.time() - t0 < wait:
        info = session_info(session)
        if info:
            try:
                s = _connect(info, 5)
                send_msg(s, {"token": info["token"], "op": "ping"})
                if recv_msg(s):
                    s.close()
                    return info
            except OSError:
                pass
        time.sleep(0.15)
    return None


def start_server(session, idle_hours=3.0, wait=20.0):
    if not _acquire_spawn_lock(session):
        # someone else is starting this session right now: wait for theirs instead of racing
        info = _wait_for_kernel(session, wait)
        if info:
            return info
        raise RuntimeError(f"another process is starting session '{session}' and it did not come up in {wait:.0f}s")
    try:
        log = open(LOGS / f"{session}.log", "a", encoding="utf-8")
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
        subprocess.Popen(
            [sys.executable, "-m", "xl.server", session, str(idle_hours)],
            cwd=str(TOOLS_DIR), stdin=subprocess.DEVNULL, stdout=log, stderr=log,
            creationflags=flags, close_fds=True,
        )
        info = _wait_for_kernel(session, wait)
        if info:
            return info
        raise RuntimeError(f"xl server for session '{session}' did not come up; see {LOGS / (session + '.log')}")
    finally:
        _release_spawn_lock(session)


def ensure(session, spawn=True, busy_retries=3):
    info = session_info(session)
    if info:
        for attempt in range(busy_retries + 1):
            try:
                s = _connect(info, 5)
                s.close()
                return info
            except OSError:
                if not _kernel_alive(info):
                    break  # dead or recycled PID: fall through to respawn
                if attempt < busy_retries:
                    time.sleep(0.5 * (attempt + 1))
        else:
            # alive but unreachable after retries: busy in a long COM call, or wedged. Never
            # replace it (that orphans the kernel and its hidden Excel); tell the caller.
            raise RuntimeError(
                f"kernel '{session}' (pid {info.get('pid')}) is alive but not accepting connections: "
                f"busy in a long call, or wedged. Retry with a longer --timeout, or 'xl stop --force'.")
    if not spawn:
        return None
    f = SESSIONS / f"{session}.json"
    if f.exists():
        try:
            f.unlink()
        except OSError:
            pass
    return start_server(session)


def request(session, payload, timeout=120.0, spawn=True):
    try:
        info = ensure(session, spawn=spawn)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    if info is None:
        return {"ok": False, "error": f"no live session '{session}'"}
    s = _connect(info, timeout)
    try:
        payload = dict(payload, token=info["token"], deadline=time.time() + timeout)
        send_msg(s, payload)
        try:
            resp = recv_msg(s)
        except socket.timeout:
            return {"ok": False, "timeout": True,
                    "error": f"no reply within {timeout:.0f}s; the kernel is still running that code "
                             f"(it will finish, but the result is lost). Use a longer --timeout, "
                             f"or 'xl stop --force' if it is stuck."}
        return resp or {"ok": False, "error": "connection closed by kernel"}
    finally:
        s.close()


def ping(session, timeout=10.0):
    return request(session, {"op": "ping"}, timeout=timeout, spawn=False)


def exec_code(code, session="default", timeout=120.0, max_chars=20000):
    return request(session, {"op": "exec", "code": code, "max_chars": max_chars}, timeout=timeout)


def _kill(pid):
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    for _ in range(50):
        if not pid_alive(pid):
            return True
        time.sleep(0.1)
    return not pid_alive(pid)


def stop(session, force=False):
    info = session_info(session)
    if not info:
        return f"{session}: no session file"
    msg = []
    try:
        s = _connect(info, 15)
        send_msg(s, {"token": info["token"], "op": "shutdown"})
        recv_msg(s)
        s.close()
        for _ in range(50):
            if not pid_alive(info["pid"]):
                break
            time.sleep(0.1)
        msg.append(f"{session}: shutdown requested")
    except OSError as e:
        msg.append(f"{session}: could not reach kernel ({e})")
    if pid_alive(info.get("pid")):
        if force:
            # Excel first (the kernel's finally never runs after a hard kill), then the kernel.
            # Both only if the PID still belongs to the process the session file describes.
            ep, es = info.get("excel_pid"), info.get("excel_started")
            if ep:
                if is_ours(ep, "EXCEL.EXE", es if es is not None else info.get("started")):
                    from .excelcom import windows_of_pid
                    if any(w["visible"] and w["class"] == "XLMAIN" for w in windows_of_pid(ep)):
                        msg.append(f"{session}: excel_pid {ep} has a visible window; refusing to kill it")
                    else:
                        msg.append(f"{session}: {'killed' if _kill(ep) else 'could not kill'} excel_pid {ep}")
                elif pid_alive(ep):
                    msg.append(f"{session}: excel_pid {ep} is not our Excel any more (recycled PID); left alone")
            kp = info.get("pid")
            if _kernel_alive(info):
                msg.append(f"{session}: {'killed' if _kill(kp) else 'could not kill'} kernel pid {kp}")
            else:
                msg.append(f"{session}: pid {kp} is not our kernel any more (recycled PID); left alone")
        else:
            msg.append(f"{session}: still alive (pid {info['pid']}); use --force")
    f = SESSIONS / f"{session}.json"
    if not _kernel_alive(info) and f.exists():
        try:
            f.unlink()
        except OSError:
            pass
    return "\n".join(msg)
