"""Persistent kernel: one process per session, single-threaded (COM STA friendly).

Protocol: TCP on 127.0.0.1, length-prefixed JSON. Requests carry the session token.
  {"token":..., "op":"exec", "code":..., "max_chars":N, "deadline": epoch_seconds}
  {"token":..., "op":"ping"} | {"token":..., "op":"shutdown"}

Lifecycle rules (paid for on 2026-09-09):
  - the session file carries `excel_pid` and `excel_started` so `xl stop --force` can reach the
    hidden Excel, and is only deleted at exit if it still names THIS kernel (a replacement kernel's
    file must survive our exit);
  - a request is read with a timeout and a size cap BEFORE the token check, so a local process
    that connects and sends nothing cannot park the single thread forever;
  - a request whose client-side deadline has already passed is skipped, not run late (a `replace`
    that timed out on the client would otherwise run again from the backlog).
"""
import ast
import contextlib
import io
import json
import os
import secrets
import socket
import struct
import sys
import time
import traceback

from .paths import SESSIONS

MAX_OUT = 20000
MAX_REQUEST = 64 * 1024 * 1024   # a 64 MB script is not a request, it is an attack or a bug
RECV_TIMEOUT = 15.0              # a real client sends the whole request at once


def _recvall(conn, n):
    buf = bytearray()
    while len(buf) < n:
        chunk = conn.recv(min(n - len(buf), 1 << 20))
        if not chunk:
            return None
        buf += chunk
    return bytes(buf)


def recv_msg(conn, max_bytes=MAX_REQUEST):
    hdr = _recvall(conn, 4)
    if not hdr:
        return None
    (n,) = struct.unpack("!I", hdr)
    if n > max_bytes:
        raise ValueError(f"message of {n:,} bytes exceeds the {max_bytes:,} byte cap")
    body = _recvall(conn, n)
    if body is None:
        return None
    return json.loads(body.decode("utf-8"))


def send_msg(conn, obj):
    data = json.dumps(obj, default=str).encode("utf-8")
    conn.sendall(struct.pack("!I", len(data)) + data)


def _trunc(s, n):
    if s is None:
        return None
    if len(s) <= n:
        return s
    return s[:n] + f"\n... [truncated {len(s) - n} chars; narrow the query or raise --max-chars]"


class Kernel:
    def __init__(self, session):
        from .helpers import build_namespace

        self.session = session
        self.ns = build_namespace(session)
        self.count = 0

    def run(self, code, max_chars=MAX_OUT):
        out, err = io.StringIO(), io.StringIO()
        result = error = None
        t0 = time.perf_counter()
        try:
            tree = ast.parse(code, mode="exec")
            last_expr = None
            if tree.body and isinstance(tree.body[-1], ast.Expr):
                last_expr = ast.Expression(tree.body[-1].value)
                ast.copy_location(last_expr, tree.body[-1])
                tree.body = tree.body[:-1]
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                if tree.body:
                    exec(compile(tree, "<xl>", "exec"), self.ns)
                if last_expr is not None:
                    val = eval(compile(last_expr, "<xl>", "eval"), self.ns)
                    self.ns["_"] = val
                    if val is not None:
                        result = val if isinstance(val, str) else repr(val)
        except BaseException as e:  # noqa: BLE001 - a kernel must survive anything
            tb = e.__traceback__.tb_next if e.__traceback__ else None
            error = "".join(traceback.format_exception(type(e), e, tb))
        self.count += 1
        return {
            "ok": error is None,
            "stdout": _trunc(out.getvalue(), max_chars),
            "stderr": _trunc(err.getvalue(), max_chars // 4),
            "result": _trunc(result, max_chars),
            "error": _trunc(error, max_chars // 2),
            "elapsed": round(time.perf_counter() - t0, 3),
        }


def _write_session(sfile, info):
    tmp = sfile.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(info))
    os.replace(tmp, sfile)


def _session_is_ours(sfile, pid):
    try:
        return json.loads(sfile.read_text()).get("pid") == pid
    except Exception:
        return False


def serve(session, idle_hours=3.0):
    import pythoncom

    pythoncom.CoInitialize()
    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.bind(("127.0.0.1", 0))
    srv.listen(64)  # every client call costs two connections (probe + request); 5 filled during one long calc
    port = srv.getsockname()[1]
    token = secrets.token_hex(16)
    sfile = SESSIONS / f"{session}.json"
    me = os.getpid()
    info = {"session": session, "port": port, "token": token, "pid": me, "started": time.time(),
            "excel_pid": None, "excel_started": None}
    kernel = Kernel(session)
    kernel.ns["_session_file"] = sfile
    kernel.ns["_session_info"] = info
    _write_session(sfile, info)
    print(f"[xl] session={session} port={port} pid={me}", flush=True)
    srv.settimeout(30)
    last = time.time()

    def _sync_excel_pid():
        ep = kernel.ns["_excel_pid"]()
        if ep != info.get("excel_pid"):
            from .procs import process_start_time
            info["excel_pid"] = ep
            info["excel_started"] = process_start_time(ep) if ep else None
            try:
                _write_session(sfile, info)
            except Exception:
                pass

    try:
        while True:
            try:
                conn, _ = srv.accept()
            except socket.timeout:
                if time.time() - last > idle_hours * 3600:
                    print("[xl] idle timeout, exiting", flush=True)
                    break
                continue
            with conn:
                conn.settimeout(RECV_TIMEOUT)
                try:
                    req = recv_msg(conn)
                except Exception as e:  # noqa: BLE001 - slow loris, oversized header, bad JSON
                    print(f"[xl] dropped connection: {type(e).__name__}: {e}", flush=True)
                    continue
                if not req or req.get("token") != token:
                    try:
                        send_msg(conn, {"ok": False, "error": "bad token"})
                    except Exception:
                        pass
                    continue
                conn.settimeout(60)
                last = time.time()
                op = req.get("op", "exec")
                deadline = req.get("deadline")
                if op == "ping":
                    _sync_excel_pid()
                    resp = {"ok": True, "pid": me, "count": kernel.count, "started": info["started"],
                            "cache": kernel.ns["cache_info"](), "excel_pid": info["excel_pid"]}
                elif op == "shutdown":
                    try:
                        send_msg(conn, {"ok": True})
                    except Exception:
                        pass
                    break
                elif deadline is not None and time.time() > float(deadline):
                    late = time.time() - float(deadline)
                    print(f"[xl] skipped a request whose client gave up {late:.0f}s ago", flush=True)
                    resp = {"ok": False, "expired": True,
                            "error": f"request skipped: its client stopped waiting {late:.0f}s ago"}
                else:
                    resp = kernel.run(req.get("code", ""), int(req.get("max_chars") or MAX_OUT))
                    _sync_excel_pid()
                try:
                    send_msg(conn, resp)
                except Exception:
                    print("[xl] client went away before reply", flush=True)
    finally:
        try:
            kernel.ns["_shutdown"]()
        except Exception:
            traceback.print_exc()
        try:
            if _session_is_ours(sfile, me):
                sfile.unlink()
        except Exception:
            pass
        print("[xl] stopped", flush=True)
        # Skip interpreter teardown: freeing a 700k-cell openpyxl graph takes many seconds.
        os._exit(0)


if __name__ == "__main__":
    serve(sys.argv[1] if len(sys.argv) > 1 else "default",
          float(sys.argv[2]) if len(sys.argv) > 2 else 3.0)
