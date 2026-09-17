from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from app.db import Database
from app.importers.hwaseong_workbook import HwaseongWorkbookParser
from app.repositories.transport_entries import TransportEntryRepository


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
        file_sha256 = hashlib.sha256(path.read_bytes()).hexdigest()
        rows = self.parser.parse(path, report_month)
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
