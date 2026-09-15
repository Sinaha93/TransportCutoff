from pathlib import Path
from app.config import RuntimePaths


def test_runtime_paths_are_relative_to_bundle(tmp_path: Path):
    paths = RuntimePaths.from_root(tmp_path)
    assert paths.database == tmp_path / "runtime" / "data" / "app.db"
    assert paths.outputs == tmp_path / "runtime" / "outputs"
    assert paths.imports == tmp_path / "runtime" / "imports"
    assert paths.backups == tmp_path / "runtime" / "backups"
    assert paths.template == tmp_path / "assets" / "report_template.pptx"


def test_runtime_paths_ensure_creates_directories_idempotently(tmp_path: Path):
    paths = RuntimePaths.from_root(tmp_path)

    paths.ensure()
    paths.ensure()

    for directory in (
        paths.database.parent,
        paths.outputs,
        paths.imports,
        paths.backups,
    ):
        assert directory.is_dir()
