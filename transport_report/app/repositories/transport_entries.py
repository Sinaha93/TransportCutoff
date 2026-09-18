from __future__ import annotations

import re
import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import date, timedelta

from app.db import Database
from app.importers.hwaseong_workbook import (
    ParsedTransportEntry,
    WorkbookStructureError,
    canonical_trip_count,
)
from app.repositories.monthly_inputs import MonthLockGuard


SOURCE_TYPE = "transport"
_SHA256 = re.compile(r"^[0-9a-fA-F]{64}$")
_SELECTABLE_STATUSES = frozenset({"imported", "imported_with_errors"})


@dataclass(frozen=True, slots=True)
class ImportCommitResult:
    batch_id: int
    status: str
    inserted_count: int
    blocking_errors: tuple[str, ...]
    is_current: bool


@dataclass(frozen=True, slots=True)
class ImportBatchSnapshot:
    """Persisted batch-selection state for report validation adapters."""

    batch_id: int
    report_month: str
    source_type: str
    source_filename: str
    file_sha256: str
    status: str
    is_current: bool

    def __post_init__(self) -> None:
        if (
            isinstance(self.batch_id, bool)
            or not isinstance(self.batch_id, int)
            or self.batch_id < 1
        ):
            raise ValueError("batch_id must be a positive integer")
        _validate_report_month(self.report_month)
        for field_name in ("source_type", "source_filename"):
            value = getattr(self, field_name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"{field_name} must be nonblank text")
        object.__setattr__(
            self, "file_sha256", _canonical_sha256(self.file_sha256)
        )
        if (
            not isinstance(self.status, str)
            or self.status not in _SELECTABLE_STATUSES
        ):
            raise ValueError("status must be imported or imported_with_errors")
        if not isinstance(self.is_current, bool):
            raise TypeError("is_current must be bool")


class TransportEntryRepository:
    def __init__(self, database: Database) -> None:
        self.database = database
        self.lock_guard = MonthLockGuard()

    def import_entries(
        self,
        *,
        report_month: str,
        source_filename: str,
        file_sha256: str,
        rows: list[ParsedTransportEntry],
    ) -> ImportCommitResult:
        _validate_import_rows(report_month, rows)
        file_sha256 = _canonical_sha256(file_sha256)
        connection = self.database.connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            duplicate = connection.execute(
                "SELECT id, error_summary, is_current FROM import_batches "
                "WHERE report_month = ? AND source_type = ? "
                "AND lower(file_sha256) = ? "
                "ORDER BY is_current DESC, id DESC LIMIT 1",
                (report_month, SOURCE_TYPE, file_sha256),
            ).fetchone()
            if duplicate is not None:
                unknown_aliases: list[str] = []
                for entry in connection.execute(
                    "SELECT unresolved_alias FROM transport_entries "
                    "WHERE import_batch_id = ? AND destination_id IS NULL "
                    "ORDER BY id",
                    (duplicate["id"],),
                ):
                    alias = entry["unresolved_alias"]
                    if alias not in unknown_aliases:
                        unknown_aliases.append(alias)
                connection.rollback()
                return ImportCommitResult(
                    batch_id=int(duplicate["id"]),
                    status="duplicate",
                    inserted_count=0,
                    blocking_errors=tuple(
                        f"Unknown destination alias: {alias}"
                        for alias in unknown_aliases
                    ),
                    is_current=_database_bool(
                        duplicate["is_current"], "import_batches.is_current"
                    ),
                )

            self.lock_guard.require_unlocked(connection, report_month)
            staged_rows, blocking_errors = self._resolve_rows(connection, rows)
            status = "imported_with_errors" if blocking_errors else "imported"
            error_summary = "\n".join(blocking_errors) or None
            batch_id = connection.execute(
                "INSERT INTO import_batches"
                "(report_month, source_type, source_filename, file_sha256, status, "
                "error_summary) VALUES (?, ?, ?, ?, 'staged', ?)",
                (
                    report_month,
                    SOURCE_TYPE,
                    source_filename,
                    file_sha256,
                    error_summary,
                ),
            ).lastrowid
            if batch_id is None:
                raise RuntimeError("Import batch insert did not return an id")

            connection.executemany(
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, unresolved_alias, "
                "source_alias, source_sheet, source_row, source_date, transport_day, "
                "transport_type, vehicle_type, vehicle_driver_group, trip_count_text, "
                "unit_rate_won, cost_won, source_note) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                [
                    (
                        batch_id,
                        report_month,
                        destination_id,
                        unresolved_alias,
                        row.destination_alias,
                        row.source_sheet,
                        row.source_row,
                        row.source_date.isoformat(),
                        row.day,
                        row.transport_type,
                        row.vehicle_type,
                        row.vehicle_driver_group,
                        canonical_trip_count(row.trip_count),
                        row.unit_rate_won,
                        row.cost_won,
                        row.source_note,
                    )
                    for row, destination_id, unresolved_alias in staged_rows
                ],
            )
            connection.execute(
                "UPDATE import_batches SET status = ? WHERE id = ?",
                (status, batch_id),
            )
            connection.execute(
                "UPDATE import_batches SET is_current = 0 "
                "WHERE report_month = ? AND source_type = ? AND is_current = 1",
                (report_month, SOURCE_TYPE),
            )
            connection.execute(
                "UPDATE import_batches SET is_current = 1 WHERE id = ?",
                (batch_id,),
            )
            connection.commit()
            return ImportCommitResult(
                batch_id=int(batch_id),
                status=status,
                inserted_count=len(staged_rows),
                blocking_errors=blocking_errors,
                is_current=True,
            )
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def list_current_batches(
        self, report_month: str
    ) -> tuple[ImportBatchSnapshot, ...]:
        """Return only persisted current batches used by downstream calculations.

        Superseded batches remain audit history and must not be included in report
        calculations or validation context construction.
        """
        _validate_report_month(report_month)
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT id, report_month, source_type, source_filename, "
                "file_sha256, status, is_current FROM import_batches "
                "WHERE report_month = ? AND is_current = 1 "
                "ORDER BY source_type, id",
                (report_month,),
            ).fetchall()
        return tuple(
            ImportBatchSnapshot(
                batch_id=int(row["id"]),
                report_month=str(row["report_month"]),
                source_type=str(row["source_type"]),
                source_filename=str(row["source_filename"]),
                file_sha256=str(row["file_sha256"]),
                status=str(row["status"]),
                is_current=_database_bool(
                    row["is_current"], "import_batches.is_current"
                ),
            )
            for row in rows
        )

    @staticmethod
    def _resolve_rows(
        connection: sqlite3.Connection, rows: list[ParsedTransportEntry]
    ) -> tuple[
        list[tuple[ParsedTransportEntry, int | None, str | None]], tuple[str, ...]
    ]:
        staged: list[tuple[ParsedTransportEntry, int | None, str | None]] = []
        unknown_aliases: list[str] = []
        for row in rows:
            alias = unicodedata.normalize("NFKC", row.destination_alias).strip()
            destination = connection.execute(
                "SELECT destination_id FROM destination_aliases "
                "WHERE raw_name = ? AND source_type = ?",
                (alias, SOURCE_TYPE),
            ).fetchone()
            if destination is None:
                staged.append((row, None, alias))
                if alias not in unknown_aliases:
                    unknown_aliases.append(alias)
            else:
                staged.append((row, int(destination["destination_id"]), None))
        blocking_errors = tuple(
            f"Unknown destination alias: {alias}" for alias in unknown_aliases
        )
        return staged, blocking_errors


def _validate_import_rows(
    report_month: str, rows: list[ParsedTransportEntry]
) -> None:
    report_month_start = _validate_report_month(report_month)

    report_year_month = (report_month_start.year, report_month_start.month)
    previous_month_day = report_month_start - timedelta(days=1)
    previous_year_month = (previous_month_day.year, previous_month_day.month)
    for row in rows:
        label = f"{row.source_sheet} row {row.source_row}"
        if row.report_month != report_month:
            raise WorkbookStructureError(
                f"{label} report_month must match import report_month "
                f"{report_month}"
            )
        if not isinstance(row.source_date, date):
            raise WorkbookStructureError(f"{label} source_date is required")
        if row.day != row.source_date.day:
            raise WorkbookStructureError(
                f"{label} day must match source_date day {row.source_date.day}"
            )
        if row.transport_type not in {"regular", "nonregular"}:
            raise WorkbookStructureError(
                f"{label} transport_type must be regular or nonregular"
            )

        source_year_month = (row.source_date.year, row.source_date.month)
        if row.transport_type == "regular" and source_year_month != report_year_month:
            raise WorkbookStructureError(
                f"{label} regular source_date must be in report_month {report_month}"
            )
        if row.transport_type == "nonregular" and source_year_month not in {
            report_year_month,
            previous_year_month,
        }:
            raise WorkbookStructureError(
                f"{label} nonregular source_date must be in report_month "
                "or the previous calendar month"
            )


def _validate_report_month(report_month: str) -> date:
    try:
        report_month_start = date.fromisoformat(f"{report_month}-01")
    except (TypeError, ValueError) as error:
        raise WorkbookStructureError("report_month must use YYYY-MM") from error
    if report_month_start.strftime("%Y-%m") != report_month:
        raise WorkbookStructureError("report_month must use YYYY-MM")
    return report_month_start


def _database_bool(value: object, field: str) -> bool:
    if type(value) is not int or value not in (0, 1):
        raise ValueError(f"{field} must be stored as integer 0 or 1")
    return bool(value)


def _canonical_sha256(value: object) -> str:
    if not isinstance(value, str) or _SHA256.fullmatch(value) is None:
        raise ValueError("file_sha256 must be 64 hexadecimal characters")
    return value.lower()
