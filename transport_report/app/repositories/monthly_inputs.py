from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from decimal import Decimal

from app.db import Database
from app.domain.models import Destination
from app.importers.protocols import QuantityRecord


_MONTH = re.compile(r"^[0-9]{4}-(0[1-9]|1[0-2])$")
_DEFAULT_QUANTITY_SOURCE = "ERP 수기 확인"
# The report's business timezone is fixed KST (UTC+09:00), with no DST rules.
_SEOUL = timezone(timedelta(hours=9), name="KST")


class MonthlyInputError(Exception):
    """Base error for monthly input operations."""


class MonthlyInputValidationError(MonthlyInputError):
    """Raised when a monthly input value is invalid."""


class MonthlyInputRevisionError(MonthlyInputError):
    """Raised before writes when the submitted form no longer matches the DB."""


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
class MonthlyFormSnapshot:
    destinations: tuple[Destination, ...]
    plans: tuple[MonthlyPlan, ...]
    actuals: tuple[MonthlyActualQuantity, ...]
    sales: MonthlySales | None
    monthly_revision: str


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


class MonthLockGuard:
    """Read current month-lock state using a caller-owned connection."""

    @staticmethod
    def is_locked(connection: sqlite3.Connection, report_month: str) -> bool:
        row = connection.execute(
            "SELECT is_locked FROM month_locks WHERE report_month = ?",
            (report_month,),
        ).fetchone()
        return row is not None and bool(row["is_locked"])

    def require_unlocked(
        self, connection: sqlite3.Connection, report_month: str
    ) -> None:
        if self.is_locked(connection, report_month):
            raise MonthLockedError(
                f"Report month {report_month} is finalized and locked"
            )


class MonthlyInputRepository:
    def __init__(self, database: Database) -> None:
        self.database = database
        self.lock_guard = MonthLockGuard()

    def get_form_snapshot(self, report_month: str) -> MonthlyFormSnapshot:
        """Read both displayed values and their revision from one SQLite snapshot."""
        _validate_month(report_month)
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN")
            state = _monthly_form_state(connection, report_month)
            destinations = tuple(
                Destination(
                    id=row["id"], name=row["name"], display_order=row["display_order"],
                    active=bool(row["active"]), required_for_report=bool(row["required_for_report"]),
                    representative_item=row["representative_item"],
                    include_quantity_total=bool(row["include_quantity_total"]),
                    include_cost_total=bool(row["include_cost_total"]),
                    include_sales_total=bool(row["include_sales_total"]),
                )
                for row in state["destinations"]
                if row["active"] and row["required_for_report"]
            )
            return MonthlyFormSnapshot(
                destinations=destinations,
                plans=tuple(_plan_from_row(row) for row in state["plans"]),
                actuals=tuple(_actual_quantity_from_row(row) for row in state["actuals"]),
                sales=_sales_from_row(state["sales"][0]) if state["sales"] else None,
                monthly_revision=_monthly_form_revision(report_month, state),
            )

    def save_form(
        self, report_month: str, rows: list[dict], sales: MonthlySales | None,
        *, expected_revision: str,
    ) -> None:
        """Save the entire required-destination form in one locked transaction.

        Missing plan/actual values clear their row; explicit zero remains data.
        Recheck the required set inside the write transaction to reject stale forms.
        """
        _validate_month(report_month)
        ids = [row["destination_id"] for row in rows]
        if len(ids) != len(set(ids)):
            raise MonthlyInputValidationError("납품처 입력 행이 중복되었습니다.")
        for row in rows:
            _validate_id(row["destination_id"])
            if (row["plan_quantity"] is None) != (row["plan_cost"] is None):
                raise MonthlyInputValidationError("계획 수량과 비용을 함께 입력하세요.")
            if row["plan_quantity"] is not None:
                _canonical_quantity(row["plan_quantity"])
                _validate_money(row["plan_cost"], "plan_cost")
            if row["actual_quantity"] is not None:
                _canonical_quantity(row["actual_quantity"])
            _optional_text(row["representative_item"], "representative_item")
            _optional_text(row["source_note"], "source_note")
        if sales is not None:
            if sales.report_month != report_month:
                raise MonthlyInputValidationError("매출의 보고 월이 다릅니다.")
            _validate_money(sales.amount_won, "sales")
            confirmed_at = _normalize_confirmed_at(sales.confirmed_at)
            _optional_text(sales.source_note, "source_note")
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN IMMEDIATE")
            self.lock_guard.require_unlocked(connection, report_month)
            current_revision = _monthly_form_revision(
                report_month, _monthly_form_state(connection, report_month)
            )
            if expected_revision != current_revision:
                raise MonthlyInputRevisionError(
                    "다른 화면에서 월 입력 또는 납품처 기준정보가 변경되었습니다. "
                    "저장하지 않았습니다. 월 입력 화면을 다시 열어 최신 값을 검토한 뒤 저장하세요."
                )
            required = {row[0] for row in connection.execute(
                "SELECT id FROM destinations WHERE active = 1 AND required_for_report = 1"
            )}
            if set(ids) != required:
                raise MonthlyInputValidationError("납품처 입력 행이 변경되었습니다. 화면을 다시 여세요.")
            for row in rows:
                destination_id = row["destination_id"]
                if row["plan_quantity"] is None:
                    connection.execute("DELETE FROM monthly_plans WHERE report_month = ? AND destination_id = ?", (report_month, destination_id))
                else:
                    connection.execute(
                        "INSERT INTO monthly_plans(report_month, destination_id, quantity_ea_text, cost_won, representative_item) VALUES (?, ?, ?, ?, ?) "
                        "ON CONFLICT(report_month, destination_id) DO UPDATE SET quantity_ea_text = excluded.quantity_ea_text, cost_won = excluded.cost_won, representative_item = excluded.representative_item, updated_at = datetime('now')",
                        (report_month, destination_id, _canonical_quantity(row["plan_quantity"]), row["plan_cost"], _optional_text(row["representative_item"], "representative_item")),
                    )
                if row["actual_quantity"] is None:
                    connection.execute("DELETE FROM monthly_actual_quantities WHERE report_month = ? AND destination_id = ?", (report_month, destination_id))
                else:
                    connection.execute(
                        "INSERT INTO monthly_actual_quantities(report_month, destination_id, quantity_ea_text, source_note) VALUES (?, ?, ?, ?) "
                        "ON CONFLICT(report_month, destination_id) DO UPDATE SET quantity_ea_text = excluded.quantity_ea_text, source_note = excluded.source_note, updated_at = datetime('now')",
                        (report_month, destination_id, _canonical_quantity(row["actual_quantity"]), _optional_text(row["source_note"], "source_note")),
                    )
            if sales is None:
                connection.execute("DELETE FROM monthly_sales WHERE report_month = ?", (report_month,))
            else:
                connection.execute(
                    "INSERT INTO monthly_sales(report_month, amount_won, source_note, confirmed_at) VALUES (?, ?, ?, ?) "
                    "ON CONFLICT(report_month) DO UPDATE SET amount_won = excluded.amount_won, source_note = excluded.source_note, confirmed_at = excluded.confirmed_at, updated_at = datetime('now')",
                    (report_month, sales.amount_won, sales.source_note, confirmed_at),
                )

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

    def clear_plan(self, report_month: str, destination_id: int) -> bool:
        _validate_month(report_month)
        _validate_id(destination_id)
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                self._require_destination(connection, destination_id)
                cursor = connection.execute(
                    "DELETE FROM monthly_plans "
                    "WHERE report_month = ? AND destination_id = ?",
                    (report_month, destination_id),
                )
                removed = cursor.rowcount > 0
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return removed

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

    def clear_sales(self, report_month: str) -> bool:
        _validate_month(report_month)
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                self._require_unlocked(connection, report_month)
                cursor = connection.execute(
                    "DELETE FROM monthly_sales WHERE report_month = ?",
                    (report_month,),
                )
                removed = cursor.rowcount > 0
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return removed

    def is_month_locked(self, report_month: str) -> bool:
        _validate_month(report_month)
        with self.database.connection() as connection:
            return self.lock_guard.is_locked(connection, report_month)

    def finalize_month(self, report_month: str, input_revision: str) -> ReportRun:
        _validate_month(report_month)
        revision = _required_opaque_text(input_revision, "input_revision")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                if self.lock_guard.is_locked(connection, report_month):
                    raise MonthLockError(
                        f"Report month {report_month} is already locked"
                    )
                locked_at = _utc_now()
                connection.execute(
                    "INSERT INTO month_locks"
                    "(report_month, is_locked, locked_at, unlocked_at, "
                    "input_revision, last_unlock_reason) VALUES (?, 1, ?, NULL, ?, NULL) "
                    "ON CONFLICT(report_month) DO UPDATE SET "
                    "is_locked = 1, locked_at = excluded.locked_at, "
                    "unlocked_at = NULL, input_revision = excluded.input_revision, "
                    "last_unlock_reason = NULL",
                    (report_month, locked_at, revision),
                )
                row = connection.execute(
                    "INSERT INTO report_runs"
                    "(report_month, status, is_locked, locked_at, input_revision) "
                    "VALUES (?, 'month_locked', 1, ?, ?) RETURNING *",
                    (report_month, locked_at, revision),
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
                    "SELECT input_revision FROM month_locks "
                    "WHERE report_month = ? AND is_locked = 1",
                    (report_month,),
                ).fetchone()
                if locked_row is None:
                    raise MonthLockError(
                        f"Report month {report_month} is not locked"
                    )
                unlocked_at = _utc_now()
                connection.execute(
                    "UPDATE month_locks SET is_locked = 0, unlocked_at = ?, "
                    "last_unlock_reason = ? WHERE report_month = ? AND is_locked = 1",
                    (unlocked_at, clean_reason, report_month),
                )
                row = connection.execute(
                    "INSERT INTO report_runs"
                    "(report_month, status, is_locked, unlocked_at, unlock_reason, "
                    "input_revision) VALUES (?, 'month_unlocked', 0, ?, ?, ?) "
                    "RETURNING *",
                    (
                        report_month,
                        unlocked_at,
                        clean_reason,
                        str(locked_row["input_revision"]),
                    ),
                ).fetchone()
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return _report_run_from_row(row)

    def _require_unlocked(
        self, connection: sqlite3.Connection, report_month: str
    ) -> None:
        self.lock_guard.require_unlocked(connection, report_month)

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
    """Provide one manual ERP aggregate per stored month/destination row."""

    def __init__(self, repository: MonthlyInputRepository) -> None:
        self.repository = repository

    def load(self, report_month: str) -> list[QuantityRecord]:
        return [
            QuantityRecord(
                report_month=record.report_month,
                destination_id=record.destination_id,
                quantity_ea=record.quantity_ea,
                source="manual_erp",
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


def _monthly_form_state(
    connection: sqlite3.Connection, report_month: str
) -> dict[str, tuple[sqlite3.Row, ...]]:
    """Only semantic form/master values, with stable ordering and explicit nulls.

    All destination masters participate so adding, activating, or reordering a
    required row invalidates an open form. Timestamps and audit IDs do not affect
    form content and are deliberately excluded from the revision.
    """
    return {
        "destinations": tuple(connection.execute(
            "SELECT id, name, display_order, active, required_for_report, "
            "representative_item, include_quantity_total, include_cost_total, "
            "include_sales_total FROM destinations ORDER BY display_order, id"
        )),
        "plans": tuple(connection.execute(
            "SELECT report_month, destination_id, quantity_ea_text, cost_won, "
            "representative_item FROM monthly_plans WHERE report_month = ? "
            "ORDER BY destination_id", (report_month,),
        )),
        "actuals": tuple(connection.execute(
            "SELECT report_month, destination_id, quantity_ea_text, source_note "
            "FROM monthly_actual_quantities WHERE report_month = ? "
            "ORDER BY destination_id", (report_month,),
        )),
        "sales": tuple(connection.execute(
            "SELECT report_month, amount_won, source_note, confirmed_at "
            "FROM monthly_sales WHERE report_month = ?", (report_month,),
        )),
    }


def _monthly_form_revision(
    report_month: str, state: dict[str, tuple[sqlite3.Row, ...]]
) -> str:
    payload = {"report_month": report_month, "values": {
        name: [tuple(row) for row in rows] for name, rows in state.items()
    }}
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _validate_id(value: object) -> None:
    _validate_sqlite_integer(value, "destination_id", minimum=1)


def _validate_sqlite_integer(value: object, field: str, *, minimum: int) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < minimum
        or value > 2**63 - 1
    ):
        qualifier = "positive" if minimum == 1 else "nonnegative"
        raise MonthlyInputValidationError(
            f"{field} must be a {qualifier} integer within SQLite range"
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
    try:
        _validate_sqlite_integer(value, field, minimum=0)
    except MonthlyInputValidationError as error:
        raise MonthlyInputValidationError(
            f"{field} must be a nonnegative integer won amount"
        ) from error


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


def _required_opaque_text(value: object, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise MonthlyInputValidationError(f"{field} must be nonblank text")
    return value


def _normalize_confirmed_at(value: object) -> str:
    if isinstance(value, str) and re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
        try:
            parsed = datetime.fromisoformat(value).replace(tzinfo=_SEOUL)
        except ValueError as error:
            raise MonthlyInputValidationError(
                "confirmed_at must be a valid ISO date or timezone-aware date/time"
            ) from error
    elif isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError as error:
            raise MonthlyInputValidationError(
                "confirmed_at must be a valid ISO date or timezone-aware date/time"
            ) from error
    elif isinstance(value, datetime):
        parsed = value
    else:
        raise MonthlyInputValidationError(
            "confirmed_at must be an ISO date or timezone-aware date/time"
        )
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise MonthlyInputValidationError(
            "confirmed_at date/time requires an explicit timezone"
        )
    return parsed.astimezone(_SEOUL).replace(microsecond=0).isoformat(
        timespec="seconds"
    )


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
        input_revision=(
            None if row["input_revision"] is None else str(row["input_revision"])
        ),
    )
