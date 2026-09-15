from pathlib import Path
from app.config import RuntimePaths


def test_runtime_paths_are_relative_to_bundle(tmp_path: Path):
    paths = RuntimePaths.from_root(tmp_path)
    assert paths.database == tmp_path / "runtime" / "data" / "app.db"
    assert paths.outputs == tmp_path / "runtime" / "outputs"
    assert paths.template == tmp_path / "assets" / "report_template.pptx"
