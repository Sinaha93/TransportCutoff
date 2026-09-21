from __future__ import annotations

import socket
from pathlib import Path

from app.launcher import (
    discover_root,
    find_available_port,
    open_existing_instance,
    prepare_runtime,
    write_active_instance,
)


def test_discover_root_supports_source_and_pyinstaller(tmp_path: Path):
    source_launcher = tmp_path / "source" / "app" / "launcher.py"
    executable = tmp_path / "portable" / "TransportReport.exe"

    assert discover_root(frozen=False, module_file=source_launcher) == tmp_path / "source"
    assert discover_root(frozen=True, executable=executable) == tmp_path / "portable"


def test_prepare_runtime_creates_portable_directories(tmp_path: Path):
    paths = prepare_runtime(tmp_path)

    assert paths.database.parent.is_dir()
    assert paths.imports.is_dir()
    assert paths.backups.is_dir()
    assert paths.outputs.is_dir()


def test_find_available_port_starts_at_8765():
    with socket.socket() as occupied:
        occupied.bind(("127.0.0.1", 8765))
        assert find_available_port() == 8766


def test_second_launch_opens_existing_local_instance(tmp_path: Path):
    runtime = tmp_path / "runtime"
    write_active_instance(runtime, "http://127.0.0.1:8765/", 4321)
    opened: list[str] = []

    found = open_existing_instance(
        runtime,
        process_running=lambda pid: pid == 4321,
        browser_open=opened.append,
    )

    assert found is True
    assert opened == ["http://127.0.0.1:8765/"]


def test_existing_instance_file_never_opens_non_local_url(tmp_path: Path):
    runtime = tmp_path / "runtime"
    write_active_instance(runtime, "http://example.com:8765/", 4321)
    opened: list[str] = []

    assert not open_existing_instance(
        runtime,
        process_running=lambda _pid: True,
        browser_open=opened.append,
    )
    assert opened == []
