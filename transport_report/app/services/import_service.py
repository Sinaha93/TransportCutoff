from __future__ import annotations

import hashlib
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile, LargeZipFile, ZipFile

from app.db import Database
from app.importers.hwaseong_workbook import HwaseongWorkbookParser
from app.repositories.transport_entries import TransportEntryRepository


MAX_INPUT_BYTES = 64 * 1024 * 1024
MAX_ZIP_MEMBERS = 2_048
MAX_ZIP_MEMBER_BYTES = 64 * 1024 * 1024
MAX_ZIP_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
_COPY_CHUNK_BYTES = 1024 * 1024


class TransportImportError(Exception):
    """Base error for failures before a transport batch can be staged."""


class ImportFileLimitError(TransportImportError):
    """Raised when an input or archive expansion limit is exceeded."""


class ImportArchiveError(TransportImportError):
    """Raised when an uploaded workbook is not a readable ZIP archive."""


@dataclass(frozen=True, slots=True)
class ImportResult:
    batch_id: int
    status: str
    inserted_count: int
    blocking_errors: tuple[str, ...]
    file_sha256: str


class TransportImportService:
    def __init__(
        self,
        database: Database,
        *,
        parser: HwaseongWorkbookParser | None = None,
        repository: TransportEntryRepository | None = None,
    ) -> None:
        self.parser = parser or HwaseongWorkbookParser()
        self.repository = repository or TransportEntryRepository(database)

    def import_transport(
        self, workbook_path: str | Path, report_month: str
    ) -> ImportResult:
        path = Path(workbook_path)
        with _immutable_workbook_snapshot(path) as (snapshot_path, file_sha256):
            rows = self.parser.parse(snapshot_path, report_month)
            committed = self.repository.import_entries(
                report_month=report_month,
                source_filename=path.name,
                file_sha256=file_sha256,
                rows=rows,
            )
        return ImportResult(
            batch_id=committed.batch_id,
            status=committed.status,
            inserted_count=committed.inserted_count,
            blocking_errors=committed.blocking_errors,
            file_sha256=file_sha256,
        )


@contextmanager
def _immutable_workbook_snapshot(source_path: Path) -> Iterator[tuple[Path, str]]:
    suffix = source_path.suffix or ".xlsx"
    temporary_path: Path | None = None
    try:
        digest = hashlib.sha256()
        copied_bytes = 0
        with source_path.open("rb") as source, NamedTemporaryFile(
            mode="wb",
            prefix="transport-import-",
            suffix=suffix,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
            while chunk := source.read(_COPY_CHUNK_BYTES):
                copied_bytes += len(chunk)
                if copied_bytes > MAX_INPUT_BYTES:
                    raise ImportFileLimitError(
                        f"Workbook exceeds input byte limit of {MAX_INPUT_BYTES}"
                    )
                digest.update(chunk)
                temporary.write(chunk)

        _validate_archive_limits(temporary_path)
        yield temporary_path, digest.hexdigest()
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def _validate_archive_limits(path: Path) -> None:
    try:
        with ZipFile(path, "r") as archive:
            members = archive.infolist()
            if len(members) > MAX_ZIP_MEMBERS:
                raise ImportFileLimitError(
                    f"Workbook ZIP member count exceeds {MAX_ZIP_MEMBERS}"
                )
            total_size = 0
            for member in members:
                if member.file_size > MAX_ZIP_MEMBER_BYTES:
                    raise ImportFileLimitError(
                        "Workbook ZIP member size exceeds "
                        f"{MAX_ZIP_MEMBER_BYTES}: {member.filename}"
                    )
                total_size += member.file_size
                if total_size > MAX_ZIP_UNCOMPRESSED_BYTES:
                    raise ImportFileLimitError(
                        "Workbook ZIP uncompressed size exceeds "
                        f"{MAX_ZIP_UNCOMPRESSED_BYTES}"
                    )
    except (BadZipFile, LargeZipFile) as error:
        raise ImportArchiveError("Workbook is not a valid ZIP archive") from error
