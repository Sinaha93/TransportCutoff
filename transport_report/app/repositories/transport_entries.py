from __future__ import annotations

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


@dataclass(frozen=True, slots=True)
class ImportCommitResult:
    batch_id: int
    status: str
    inserted_count: int
    blocking_errors: tuple[str, ...]


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
        connection = self.database.connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            duplicate = connection.execute(
                "SELECT id, error_summary FROM import_batches "
                "WHERE report_month = ? AND source_type = ? AND file_sha256 = ?",
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
            connection.commit()
            return ImportCommitResult(
                batch_id=int(batch_id),
                status=status,
                inserted_count=len(staged_rows),
                blocking_errors=blocking_errors,
            )
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

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
    try:
        report_month_start = date.fromisoformat(f"{report_month}-01")
    except ValueError as error:
        raise WorkbookStructureError("report_month must use YYYY-MM") from error
    if report_month_start.strftime("%Y-%m") != report_month:
        raise WorkbookStructureError("report_month must use YYYY-MM")

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
