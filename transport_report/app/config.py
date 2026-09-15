from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RuntimePaths:
    root: Path
    database: Path
    outputs: Path
    imports: Path
    backups: Path
    template: Path

    @classmethod
    def from_root(cls, root: Path) -> "RuntimePaths":
        return cls(
            root=root,
            database=root / "runtime" / "data" / "app.db",
            outputs=root / "runtime" / "outputs",
            imports=root / "runtime" / "imports",
            backups=root / "runtime" / "backups",
            template=root / "assets" / "report_template.pptx",
        )

    def ensure(self) -> None:
        for path in (self.database.parent, self.outputs, self.imports, self.backups):
            path.mkdir(parents=True, exist_ok=True)
