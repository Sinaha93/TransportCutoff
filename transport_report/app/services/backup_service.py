"""SQLite online backups and explicitly offline restores.

restore(exclusive_access=True) is a caller contract: stop the web server and
close ALL database connections (including other processes) before calling.
SQLite locking is also checked, but cannot prove absence of idle connections.
The launcher must not offer restore while its server is running.
"""
from __future__ import annotations

from dataclasses import dataclass
from contextlib import closing
from datetime import datetime
import os
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile

from app.db import Database, _sql_statements


_SQL_TOKEN = re.compile(
    r"""'(?:''|[^'])*'|"(?:""|[^"])*"|`(?:``|[^`])*`|\[[^\]]*\]"""
    r"|--[^\r\n]*|/\*[\s\S]*?\*/|[A-Za-z_][A-Za-z_0-9]*|[0-9]+|[^\s]"
)


def _application_schema(connection):
    """Compare schema definitions without root pages, statistics, or formatting.

    Table SQL contains CHECK/UNIQUE/FK constraints omitted by table_info.
    Explicit indexes, triggers and views must also match the bundled schema.
    sqlite_* objects are engine-owned (automatic indexes/statistics/sequence).
    Preserve quoted text verbatim: whitespace/case inside a CHECK literal or
    trigger error message is part of the schema, not SQL formatting.
    """
    return {
        (kind, name, table): tuple(
            token if token[0] in "'\"`[" else token.lower()
            for token in _SQL_TOKEN.findall(sql or "")
            if not token.startswith(("--", "/*"))
        )
        for kind, name, table, sql in connection.execute(
            "SELECT type, name, tbl_name, sql FROM sqlite_master "
            "WHERE name NOT GLOB 'sqlite_*'"
        )
    }


class BackupError(ValueError):
    pass


@dataclass(frozen=True)
class RestoreResult:
    pre_restore_backup: Path


class BackupService:
    def __init__(self, database: Database, directory: Path, *, clock=datetime.now):
        self.database, self.directory, self.clock = database, Path(directory), clock

    def _validate(self, path):
        if not Path(path).is_file():
            raise BackupError("복원할 SQLite 백업 파일을 선택하세요.")
        with Path(path).open("rb") as handle:
            if handle.read(16) != b"SQLite format 3\x00":
                raise BackupError("SQLite 백업 형식이 아닙니다. 정상 백업 파일을 선택하세요.")
        try:
            connection = sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)
            expected = sqlite3.connect(":memory:")
            try:
                migrations = self.database._migration_files()
                versions = {row[0] for row in connection.execute("SELECT version FROM schema_migrations")}
                if versions != {version for version, _ in migrations}:
                    raise BackupError("백업의 스키마 버전이 현재 앱과 다릅니다. 같은 버전으로 만든 백업을 선택하세요.")
                if connection.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
                    raise BackupError("백업 무결성 검사에 실패했습니다. 다른 정상 백업을 선택하세요.")
                if connection.execute("PRAGMA foreign_key_check").fetchone() is not None:
                    raise BackupError("백업의 데이터 연결이 손상되었습니다. 다른 정상 백업을 선택하세요.")
                for _, file in migrations:
                    for statement in _sql_statements(file.read_text(encoding="utf-8")):
                        expected.execute(statement)
                if _application_schema(connection) != _application_schema(expected):
                    raise BackupError("백업의 앱 스키마(테이블·제약·인덱스·트리거)가 현재 앱과 다릅니다. 이 앱에서 만든 변경되지 않은 정상 백업을 선택하세요.")
                return migrations[-1][0]
            finally:
                connection.close()
                expected.close()
        except sqlite3.DatabaseError as error:
            raise BackupError("백업 파일이 손상되었거나 앱 데이터가 아닙니다. 정상 백업을 선택하세요.") from error

    @staticmethod
    def _temporary(parent, prefix):
        descriptor, filename = tempfile.mkstemp(prefix=prefix, suffix=".db", dir=parent)
        os.close(descriptor)
        return Path(filename)

    @staticmethod
    def _copy(source, destination):
        with closing(sqlite3.connect(Path(source).resolve().as_uri() + "?mode=ro", uri=True)) as original:
            target = sqlite3.connect(destination)
            try:
                original.backup(target)
                target.execute("PRAGMA journal_mode=DELETE")
            finally:
                target.close()

    def backup(self, *, prefix="backup"):
        self.directory.mkdir(parents=True, exist_ok=True)
        temporary = self._temporary(self.directory, ".backup-")
        try:
            self._copy(self.database.path, temporary)
            version = self._validate(temporary)
            base = f"{prefix}-{self.clock():%Y%m%d-%H%M%S}-v{version:03}"
            index = 0
            while True:
                final = self.directory / (base + (f"-{index:03}" if index else "") + ".db")
                # Hard-link publication is atomic and refuses any existing name
                # on Windows and POSIX; source and target are siblings.
                try:
                    os.link(temporary, final)
                    return final
                except FileExistsError:
                    index += 1
        except Exception as error:
            if isinstance(error, BackupError):
                raise
            raise BackupError("백업을 저장하지 못했습니다. 저장 공간·파일 사용 여부를 확인하세요.") from error
        finally:
            temporary.unlink(missing_ok=True)

    def restore(self, candidate: Path, *, exclusive_access=False):
        if exclusive_access is not True:
            raise BackupError("복원 전에 앱 서버와 모든 데이터베이스 연결을 종료하고 독점 사용을 확인하세요.")
        candidate = Path(candidate)
        try:
            self._validate(candidate)
        except OSError as error:
            raise BackupError("백업 파일을 읽지 못했습니다. 파일 사용 여부를 확인하세요.") from error
        live = self.database.path
        temporary = self._temporary(live.parent, ".restore-")
        rollback = self._temporary(live.parent, ".rollback-")
        replaced = False
        try:
            # Normalize the candidate with SQLite, excluding stale journal files.
            self._copy(candidate, temporary)
            self._validate(temporary)
            guard = sqlite3.connect(live, timeout=0)
            try:
                busy, _, _ = guard.execute("PRAGMA wal_checkpoint(TRUNCATE)").fetchone()
                if busy:
                    raise BackupError("사용 중인 데이터베이스 연결이 있습니다. 모든 연결을 종료하세요.")
                guard.execute("BEGIN EXCLUSIVE")
                guard.rollback()
            finally:
                guard.close()
            before = self.backup(prefix="pre-restore")
            shutil.copy2(live, rollback)
            os.replace(temporary, live)
            replaced = True
            self._verify_restored()
            return RestoreResult(before)
        except Exception as error:
            if replaced:
                os.replace(rollback, live)
            if isinstance(error, BackupError):
                raise
            raise BackupError("복원을 완료하지 못했습니다. 기존 데이터와 복원 전 백업은 보존되었습니다. 모든 연결·파일 사용 여부를 확인하세요.") from error
        finally:
            temporary.unlink(missing_ok=True)
            rollback.unlink(missing_ok=True)

    def _verify_restored(self):
        self._validate(self.database.path)
