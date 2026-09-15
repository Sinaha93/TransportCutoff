from __future__ import annotations

import re
import sqlite3
from pathlib import Path


_MIGRATION_NAME = re.compile(r"^(?P<version>\d+)_.*\.sql$")


class Database:
    def __init__(self, path: Path, migrations_dir: Path | None = None) -> None:
        self.path = Path(path)
        self.migrations_dir = migrations_dir or Path(__file__).with_name("db") / "migrations"

    def connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def migrate(self) -> None:
        migrations = self._migration_files()
        connection = self.connect()
        try:
            for version, migration_path in migrations:
                sql = migration_path.read_text(encoding="utf-8")
                try:
                    connection.execute("BEGIN IMMEDIATE")
                    if version in self._applied_versions(connection):
                        connection.commit()
                        continue
                    for statement in _sql_statements(sql):
                        connection.execute(statement)
                    connection.execute(
                        "INSERT INTO schema_migrations(version, applied_at) "
                        "VALUES (?, datetime('now'))",
                        (version,),
                    )
                    connection.commit()
                except Exception:
                    connection.rollback()
                    raise
        finally:
            connection.close()

    def table_names(self) -> set[str]:
        connection = self.connect()
        try:
            rows = connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
            return {str(row["name"]) for row in rows}
        finally:
            connection.close()

    def _migration_files(self) -> list[tuple[int, Path]]:
        if not self.migrations_dir.is_dir():
            raise FileNotFoundError(
                f"Migration directory does not exist: {self.migrations_dir}"
            )
        sql_files = list(self.migrations_dir.glob("*.sql"))
        if not sql_files:
            raise RuntimeError(f"No migrations found in: {self.migrations_dir}")

        migrations: list[tuple[int, Path]] = []
        seen_versions: set[int] = set()
        for path in sql_files:
            match = _MIGRATION_NAME.fullmatch(path.name)
            if match is None:
                raise ValueError(f"Invalid migration filename: {path.name}")
            version = int(match.group("version"))
            if version in seen_versions:
                raise ValueError(f"Duplicate migration version: {version}")
            seen_versions.add(version)
            migrations.append((version, path))
        return sorted(migrations)

    @staticmethod
    def _applied_versions(connection: sqlite3.Connection) -> set[int]:
        exists = connection.execute(
            "SELECT 1 FROM sqlite_master "
            "WHERE type = 'table' AND name = 'schema_migrations'"
        ).fetchone()
        if exists is None:
            return set()
        return {
            int(row["version"])
            for row in connection.execute("SELECT version FROM schema_migrations")
        }


def _sql_statements(script: str) -> list[str]:
    statements: list[str] = []
    current: list[str] = []
    for character in script:
        current.append(character)
        if character == ";" and sqlite3.complete_statement("".join(current)):
            statement = "".join(current).strip()
            if statement:
                statements.append(statement)
            current.clear()
    trailing = "".join(current).strip()
    if trailing:
        statements.append(trailing)
    return statements
