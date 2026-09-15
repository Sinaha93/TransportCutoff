from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal

from app.db import Database
from app.importers.protocols import QuantityRecord


_MONTH = re.compile(r"^[0-9]{4}-(0[1-9]|1[0-2])$")
_DEFAULT_QUANTITY_SOURCE = "ERP 수기 확인"


class MonthlyInputError(Exception):
    """Base error for monthly input operations."""


class MonthlyInputValidationError(MonthlyInputError):
    """Raised when a monthly input value is invalid."""


class DestinationNotFoundError(MonthlyInputError):
    """Raised when a monthly input references an unknown destination."""


class MonthLockedError(MonthlyInputError):
    """Raised when an ordinary write targets a finalized month."""


class MonthLockError(MonthlyInputError):
    """Raised when a requested lock transition is invalid."""


@dataclass(frozen=True, slots=True)
class MonthlyPlan:
    report_month: str
    destination_id: int
    quantity_ea: Decimal
    cost_won: int
    representative_item: str | None


@dataclass(frozen=True, slots=True)
class MonthlyActualQuantity:
    report_month: str
    destination_id: int
    quantity_ea: Decimal
    source_note: str | None


@dataclass(frozen=True, slots=True)
class MonthlySales:
    report_month: str
    amount_won: int
    source_note: str | None
    confirmed_at: str


@dataclass(frozen=True, slots=True)
class ReportRun:
    id: int
    report_month: str
    status: str
    is_locked: bool
    locked_at: str | None
    unlocked_at: str | None
    unlock_reason: str | None
    input_revision: str | None


class MonthlyInputRepository:
    def __init__(self, database: Database) -> None:
        self.database = database

    def save_plan(
        self,
        report_month: str,
        destination_id: int,
        quantity_ea: Decimal | int,
        cost_won: int,
        representative_item: str | None,
    ) -> MonthlyPlan:
        _validate_month(report_month)
        _validate_id(destination_id)
        quantity_text = _canonical_quantity(quantity_ea)
        _validate_money(cost_won, "cost_won")
        clean_item = _optional_text(representative_item, "representative_item")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                self._require_destination(connection, destination_id)
                row = connection.execute(
                    "INSERT INTO monthly_plans"
                    "(report_month, destination_id, quantity_ea_text, cost_won, "
                    "representative_item) VALUES (?, ?, ?, ?, ?) "
                    "ON CONFLICT(report_month, destination_id) DO UPDATE SET "
                    "quantity_ea_text = excluded.quantity_ea_text, "
                    "cost_won = excluded.cost_won, "
                    "representative_item = excluded.representative_item, "
                    "updated_at = datetime('now') RETURNING *",
                    (
                        report_month,
                        destination_id,
                        quantity_text,
                        cost_won,
                        clean_item,
                    ),
                ).fetchone()
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                _raise_integrity_error(error, destination_id, "Plan")
            except Exception:
                connection.rollback()
                raise
        return _plan_from_row(row)

    def get_plan(
        self, report_month: str, destination_id: int
    ) -> MonthlyPlan | None:
        _validate_month(report_month)
        _validate_id(destination_id)
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM monthly_plans "
                "WHERE report_month = ? AND destination_id = ?",
                (report_month, destination_id),
            ).fetchone()
        return None if row is None else _plan_from_row(row)

    def list_plans(self, report_month: str) -> list[MonthlyPlan]:
        _validate_month(report_month)
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT p.* FROM monthly_plans AS p "
                "JOIN destinations AS d ON d.id = p.destination_id "
                "WHERE p.report_month = ? ORDER BY d.display_order, d.id",
                (report_month,),
            ).fetchall()
        return [_plan_from_row(row) for row in rows]

    def save_actual_quantity(
        self,
        report_month: str,
        destination_id: int,
        quantity_ea: Decimal | int,
        source_note: str | None = _DEFAULT_QUANTITY_SOURCE,
    ) -> MonthlyActualQuantity:
        _validate_month(report_month)
        _validate_id(destination_id)
        quantity_text = _canonical_quantity(quantity_ea)
        clean_source = _optional_text(source_note, "source_note")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                self._require_destination(connection, destination_id)
                row = connection.execute(
                    "INSERT INTO monthly_actual_quantities"
                    "(report_month, destination_id, quantity_ea_text, source_note) "
                    "VALUES (?, ?, ?, ?) "
                    "ON CONFLICT(report_month, destination_id) DO UPDATE SET "
                    "quantity_ea_text = excluded.quantity_ea_text, "
                    "source_note = excluded.source_note, "
                    "updated_at = datetime('now') RETURNING *",
                    (report_month, destination_id, quantity_text, clean_source),
                ).fetchone()
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                _raise_integrity_error(error, destination_id, "Actual quantity")
            except Exception:
                connection.rollback()
                raise
        return _actual_quantity_from_row(row)

    def get_actual_quantity(
        self, report_month: str, destination_id: int
    ) -> MonthlyActualQuantity | None:
        _validate_month(report_month)
        _validate_id(destination_id)
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM monthly_actual_quantities "
                "WHERE report_month = ? AND destination_id = ?",
                (report_month, destination_id),
            ).fetchone()
        return None if row is None else _actual_quantity_from_row(row)

    def list_actual_quantities(
        self, report_month: str
    ) -> list[MonthlyActualQuantity]:
        _validate_month(report_month)
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT q.* FROM monthly_actual_quantities AS q "
                "JOIN destinations AS d ON d.id = q.destination_id "
                "WHERE q.report_month = ? ORDER BY d.display_order, d.id",
                (report_month,),
            ).fetchall()
        return [_actual_quantity_from_row(row) for row in rows]

    def clear_actual_quantity(self, report_month: str, destination_id: int) -> bool:
        _validate_month(report_month)
        _validate_id(destination_id)
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                self._require_destination(connection, destination_id)
                cursor = connection.execute(
                    "DELETE FROM monthly_actual_quantities "
                    "WHERE report_month = ? AND destination_id = ?",
                    (report_month, destination_id),
                )
                removed = cursor.rowcount > 0
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return removed

    def save_sales(
        self,
        report_month: str,
        amount_won: int,
        source_note: str | None,
        confirmed_at: str | datetime,
    ) -> MonthlySales:
        _validate_month(report_month)
        _validate_money(amount_won, "amount_won")
        clean_source = _optional_text(source_note, "source_note")
        clean_timestamp = _normalize_confirmed_at(confirmed_at)
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                row = connection.execute(
                    "INSERT INTO monthly_sales"
                    "(report_month, amount_won, source_note, confirmed_at) "
                    "VALUES (?, ?, ?, ?) "
                    "ON CONFLICT(report_month) DO UPDATE SET "
                    "amount_won = excluded.amount_won, "
                    "source_note = excluded.source_note, "
                    "confirmed_at = excluded.confirmed_at, "
                    "updated_at = datetime('now') RETURNING *",
                    (report_month, amount_won, clean_source, clean_timestamp),
                ).fetchone()
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return _sales_from_row(row)

    def get_sales(self, report_month: str) -> MonthlySales | None:
        _validate_month(report_month)
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM monthly_sales WHERE report_month = ?",
                (report_month,),
            ).fetchone()
        return None if row is None else _sales_from_row(row)

    def is_month_locked(self, report_month: str) -> bool:
        _validate_month(report_month)
        with self.database.connection() as connection:
            return self._month_is_locked(connection, report_month)

    def finalize_month(self, report_month: str, input_revision: str) -> ReportRun:
        _validate_month(report_month)
        clean_revision = _required_text(input_revision, "input_revision")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                if self._month_is_locked(connection, report_month):
                    raise MonthLockError(
                        f"Report month {report_month} is already locked"
                    )
                row = connection.execute(
                    "INSERT INTO report_runs"
                    "(report_month, status, is_locked, locked_at, input_revision) "
                    "VALUES (?, 'finalized', 1, ?, ?) RETURNING *",
                    (report_month, _utc_now(), clean_revision),
                ).fetchone()
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return _report_run_from_row(row)

    def unlock_month(self, report_month: str, reason: str) -> ReportRun:
        _validate_month(report_month)
        clean_reason = _required_text(reason, "unlock reason")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                locked_row = connection.execute(
                    "SELECT id FROM report_runs "
                    "WHERE report_month = ? AND is_locked = 1 "
                    "ORDER BY id DESC LIMIT 1",
                    (report_month,),
                ).fetchone()
                if locked_row is None:
                    raise MonthLockError(
                        f"Report month {report_month} is not locked"
                    )
                row = connection.execute(
                    "UPDATE report_runs SET status = 'unlocked', is_locked = 0, "
                    "unlocked_at = ?, unlock_reason = ? WHERE id = ? RETURNING *",
                    (_utc_now(), clean_reason, int(locked_row["id"])),
                ).fetchone()
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return _report_run_from_row(row)

    @staticmethod
    def _month_is_locked(
        connection: sqlite3.Connection, report_month: str
    ) -> bool:
        return connection.execute(
            "SELECT 1 FROM report_runs "
            "WHERE report_month = ? AND is_locked = 1 LIMIT 1",
            (report_month,),
        ).fetchone() is not None

    def _require_unlocked(
        self, connection: sqlite3.Connection, report_month: str
    ) -> None:
        if self._month_is_locked(connection, report_month):
            raise MonthLockedError(f"Report month {report_month} is finalized and locked")

    @staticmethod
    def _require_destination(
        connection: sqlite3.Connection, destination_id: int
    ) -> None:
        if connection.execute(
            "SELECT 1 FROM destinations WHERE id = ?", (destination_id,)
        ).fetchone() is None:
            raise DestinationNotFoundError(
                f"Destination {destination_id} does not exist"
            )


class ManualQuantityProvider:
    def __init__(self, repository: MonthlyInputRepository) -> None:
        self.repository = repository

    def load(self, report_month: str) -> list[QuantityRecord]:
        return [
            QuantityRecord(
                report_month=record.report_month,
                destination_id=record.destination_id,
                quantity_ea=record.quantity_ea,
                source=record.source_note or "manual",
            )
            for record in self.repository.list_actual_quantities(report_month)
        ]


def _validate_month(value: object) -> None:
    if not isinstance(value, str) or _MONTH.fullmatch(value) is None:
        raise MonthlyInputValidationError("report month must use ASCII YYYY-MM")


def _raise_integrity_error(
    error: sqlite3.IntegrityError, destination_id: int, operation: str
) -> None:
    if "FOREIGN KEY constraint failed" in str(error):
        raise DestinationNotFoundError(
            f"Destination {destination_id} does not exist"
        ) from error
    raise MonthlyInputError(f"{operation} could not be saved: {error}") from error


def _validate_id(value: object) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise MonthlyInputValidationError(
            "destination_id must be a positive integer"
        )


def _canonical_quantity(value: object) -> str:
    if isinstance(value, bool) or not isinstance(value, (Decimal, int)):
        raise MonthlyInputValidationError(
            "quantity must be a nonnegative finite Decimal or integer"
        )
    decimal_value = Decimal(value)
    if not decimal_value.is_finite() or decimal_value < 0:
        raise MonthlyInputValidationError(
            "quantity must be a nonnegative finite Decimal or integer"
        )
    if decimal_value == 0:
        return "0"
    text = format(decimal_value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def _validate_money(value: object, field: str) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise MonthlyInputValidationError(
            f"{field} must be a nonnegative integer won amount"
        )


def _optional_text(value: object, field: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise MonthlyInputValidationError(f"{field} must be text or None")
    clean_value = value.strip()
    return clean_value or None


def _required_text(value: object, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise MonthlyInputValidationError(f"{field} must be nonblank text")
    return value.strip()


def _normalize_confirmed_at(value: object) -> str:
    if isinstance(value, datetime):
        return value.replace(microsecond=0).isoformat(timespec="seconds")
    if not isinstance(value, str):
        raise MonthlyInputValidationError(
            "confirmed_at must be a normalized ISO date/time string or datetime"
        )
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as error:
        raise MonthlyInputValidationError(
            "confirmed_at must be a normalized ISO date/time string or datetime"
        ) from error
    normalized = parsed.isoformat(timespec="seconds")
    if "T" not in value or value != normalized:
        raise MonthlyInputValidationError(
            "confirmed_at must be a normalized ISO date/time string or datetime"
        )
    return normalized


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat(timespec="seconds")


def _plan_from_row(row: sqlite3.Row) -> MonthlyPlan:
    return MonthlyPlan(
        report_month=str(row["report_month"]),
        destination_id=int(row["destination_id"]),
        quantity_ea=Decimal(str(row["quantity_ea_text"])),
        cost_won=int(row["cost_won"]),
        representative_item=row["representative_item"],
    )


def _actual_quantity_from_row(row: sqlite3.Row) -> MonthlyActualQuantity:
    return MonthlyActualQuantity(
        report_month=str(row["report_month"]),
        destination_id=int(row["destination_id"]),
        quantity_ea=Decimal(str(row["quantity_ea_text"])),
        source_note=row["source_note"],
    )


def _sales_from_row(row: sqlite3.Row) -> MonthlySales:
    return MonthlySales(
        report_month=str(row["report_month"]),
        amount_won=int(row["amount_won"]),
        source_note=row["source_note"],
        confirmed_at=str(row["confirmed_at"]),
    )


def _report_run_from_row(row: sqlite3.Row) -> ReportRun:
    return ReportRun(
        id=int(row["id"]),
        report_month=str(row["report_month"]),
        status=str(row["status"]),
        is_locked=bool(row["is_locked"]),
        locked_at=row["locked_at"],
        unlocked_at=row["unlocked_at"],
        unlock_reason=row["unlock_reason"],
        input_revision=row["input_revision"],
    )
