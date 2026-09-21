"""Entry point for the copy-and-run Windows bundle."""
from __future__ import annotations

import json
import os
import secrets
import socket
import sys
import threading
import time
import webbrowser
from collections.abc import Callable
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from urllib.request import urlopen

import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from starlette.background import BackgroundTask

from app.config import RuntimePaths
from app.main import create_app


HOST = "127.0.0.1"
FIRST_PORT = 8765
ACTIVE_FILE = "active-instance.json"


def discover_root(
    *,
    frozen: bool | None = None,
    executable: str | Path | None = None,
    module_file: str | Path | None = None,
) -> Path:
    """Return the folder users copy, both from source and from PyInstaller."""
    is_frozen = bool(getattr(sys, "frozen", False)) if frozen is None else frozen
    if is_frozen:
        return Path(executable or sys.executable).resolve().parent
    return Path(module_file or __file__).resolve().parents[1]


def prepare_runtime(root: Path) -> RuntimePaths:
    paths = RuntimePaths.from_root(Path(root))
    paths.ensure()
    return paths


def find_available_port(host: str = HOST, start: int = FIRST_PORT) -> int:
    for port in range(start, 65536):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as candidate:
            try:
                candidate.bind((host, port))
            except OSError:
                continue
            return port
    raise RuntimeError("사용 가능한 로컬 포트를 찾지 못했습니다.")


def _active_path(runtime: Path) -> Path:
    return Path(runtime) / ACTIVE_FILE


def write_active_instance(runtime: Path, url: str, pid: int) -> None:
    runtime.mkdir(parents=True, exist_ok=True)
    destination = _active_path(runtime)
    temporary = destination.with_name(f".{destination.name}.{os.getpid()}.tmp")
    temporary.write_text(
        json.dumps({"url": url, "pid": pid}, ensure_ascii=False),
        encoding="utf-8",
    )
    temporary.replace(destination)


def _is_local_url(url: str) -> bool:
    try:
        parsed = urlsplit(url)
        return (
            parsed.scheme == "http"
            and parsed.hostname == HOST
            and parsed.port is not None
            and parsed.username is None
            and parsed.password is None
        )
    except ValueError:
        return False


def _process_running(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except (OSError, ValueError):
        return False
    return True


def open_existing_instance(
    runtime: Path,
    *,
    process_running: Callable[[int], bool] = _process_running,
    browser_open: Callable[[str], object] = webbrowser.open,
) -> bool:
    try:
        state = json.loads(_active_path(runtime).read_text(encoding="utf-8"))
        url = state["url"]
        pid = int(state["pid"])
    except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        return False
    if not isinstance(url, str) or not _is_local_url(url) or not process_running(pid):
        return False
    browser_open(url)
    return True


def _remove_active_instance(runtime: Path, pid: int) -> None:
    path = _active_path(runtime)
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
        if int(state.get("pid", -1)) == pid:
            path.unlink(missing_ok=True)
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return


def install_shutdown_action(
    application: FastAPI,
    token: str,
    request_shutdown: Callable[[], None],
) -> None:
    application.state.shutdown_token = token

    async def shutdown(request: Request) -> HTMLResponse:
        values = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
        submitted = values.get("token", [])
        if len(submitted) != 1 or not secrets.compare_digest(submitted[0], token):
            return HTMLResponse("종료 요청 확인 정보가 올바르지 않습니다.", status_code=403)
        return HTMLResponse(
            "<!doctype html><html lang='ko'><meta charset='utf-8'>"
            "<title>종료</title><p>운반비 보고 프로그램을 종료했습니다. 이 창을 닫아도 됩니다.</p></html>",
            background=BackgroundTask(request_shutdown),
        )

    application.add_api_route("/shutdown", shutdown, methods=["POST"])


def _open_browser_when_ready(url: str) -> None:
    for _ in range(50):
        try:
            with urlopen(f"{url}health", timeout=0.2) as response:
                if response.status == 200:
                    webbrowser.open(url)
                    return
        except OSError:
            time.sleep(0.1)


def main() -> int:
    root = discover_root()
    paths = prepare_runtime(root)
    runtime = root / "runtime"
    if open_existing_instance(runtime):
        return 0

    port = find_available_port()
    url = f"http://{HOST}:{port}/"
    application = create_app(paths=paths)
    server = uvicorn.Server(
        uvicorn.Config(application, host=HOST, port=port, log_level="warning")
    )
    install_shutdown_action(
        application, secrets.token_urlsafe(32), lambda: setattr(server, "should_exit", True)
    )

    pid = os.getpid()
    write_active_instance(runtime, url, pid)
    browser_thread = threading.Thread(
        target=_open_browser_when_ready, args=(url,), daemon=True
    )
    browser_thread.start()
    try:
        server.run()
    finally:
        _remove_active_instance(runtime, pid)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
