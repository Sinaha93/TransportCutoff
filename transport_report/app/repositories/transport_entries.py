from __future__ import annotations

import sqlite3
import unicodedata
from dataclasses import dataclass
from decimal import Decimal

from app.db import Database
from app.importers.hwaseong_workbook import ParsedTransportEntry
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
        connection = self.database.connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            duplicate = connection.execute(
                "SELECT id FROM import_batches "
                "WHERE report_month = ? AND source_type = ? AND file_sha256 = ?",
                (report_month, SOURCE_TYPE, file_sha256),
            ).fetchone()
            if duplicate is not None:
                connection.rollback()
                return ImportCommitResult(
                    batch_id=int(duplicate["id"]),
                    status="duplicate",
                    inserted_count=0,
                    blocking_errors=(),
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
                "source_sheet, source_row, transport_day, transport_type, "
                "vehicle_type, vehicle_driver_group, trip_count_text, unit_rate_won, "
                "cost_won, source_note) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                [
                    (
                        batch_id,
                        report_month,
                        destination_id,
                        unresolved_alias,
                        row.source_sheet,
                        row.source_row,
                        row.day,
                        row.transport_type,
                        row.vehicle_type,
                        row.vehicle_driver_group,
                        _canonical_decimal(row.trip_count),
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


def _canonical_decimal(value: Decimal) -> str:
    if value == 0:
        return "0"
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text
