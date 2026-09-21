"""Canonical manual inputs that supplement destination plans and actual sales."""
from __future__ import annotations

import hashlib
import json
import sqlite3
from dataclasses import dataclass
from decimal import Decimal

from app.db import Database
from app.repositories.monthly_inputs import (
    MonthLockGuard, MonthlyInputValidationError, MonthlyInputRevisionError,
    _canonical_quantity, _optional_text, _validate_id, _validate_money, _validate_month,
)


@dataclass(frozen=True, slots=True)
class ReportSupplementalInputs:
    report_month: str
    planned_sales_won: int | None = None
    nonregular_planned_quantity_ea: Decimal | None = None
    nonregular_planned_cost_won: int | None = None
    nonregular_actual_quantity_ea: Decimal | None = None
    nonregular_actual_cost_won: int | None = None
    source_note: str | None = None
    updated_at: str | None = None


@dataclass(frozen=True, slots=True)
class ReportInputFormSnapshot:
    current: ReportSupplementalInputs | None
    next_inputs: ReportSupplementalInputs | None
    next_month: str
    destinations: tuple[dict, ...]
    plans: tuple[dict, ...]
    revision: str


def next_report_month(month: str) -> str:
    _validate_month(month)
    year, number = map(int, month.split("-"))
    result = f"{year + (number == 12):04d}-{number % 12 + 1:02d}"
    _validate_month(result)
    return result


class ReportInputRepository:
    def __init__(self, database: Database):
        self.database = database
        self.lock_guard = MonthLockGuard()

    def get(self, report_month: str) -> ReportSupplementalInputs | None:
        _validate_month(report_month)
        with self.database.connection() as connection:
            return _get(connection, report_month)

    def save(self, inputs: ReportSupplementalInputs) -> ReportSupplementalInputs:
        values = _values(inputs)
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN IMMEDIATE")
            self.lock_guard.require_unlocked(connection, inputs.report_month)
            _save(connection, values)
            return _get(connection, inputs.report_month)

    def get_form_snapshot(self, report_month: str) -> ReportInputFormSnapshot:
        next_month = next_report_month(report_month)
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN")
            state = _state(connection, report_month, next_month)
            return ReportInputFormSnapshot(
                _get(connection, report_month), _get(connection, next_month), next_month,
                tuple(row for row in state["destinations"] if row["active"] and row["required_for_report"]),
                tuple(state["plans"]), _revision(state),
            )

    def save_form(
        self, report_month: str, current: ReportSupplementalInputs,
        next_inputs: ReportSupplementalInputs, rows: list[dict], *, expected_revision: str,
    ) -> None:
        next_month = next_report_month(report_month)
        if current.report_month != report_month or next_inputs.report_month != next_month:
            raise MonthlyInputValidationError("보고 기준월과 다음 달 계획 월을 확인하세요.")
        current_values, next_values = _values(current), _values(next_inputs)
        ids = [row["destination_id"] for row in rows]
        if len(ids) != len(set(ids)):
            raise MonthlyInputValidationError("납품처 입력 행이 중복되었습니다.")
        validated = []
        for row in rows:
            _validate_id(row["destination_id"])
            q, c = row["plan_quantity"], row["plan_cost"]
            if (q is None) != (c is None):
                raise MonthlyInputValidationError("계획 수량과 비용을 함께 입력하세요.")
            text = None if q is None else _canonical_quantity(q)
            if c is not None:
                _validate_money(c, "계획 비용")
            validated.append((row["destination_id"], text, c, _optional_text(row["representative_item"], "대표 품목")))
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN IMMEDIATE")
            self.lock_guard.require_unlocked(connection, report_month)
            self.lock_guard.require_unlocked(connection, next_month)
            state = _state(connection, report_month, next_month)
            if expected_revision != _revision(state):
                raise MonthlyInputRevisionError("다른 화면에서 보고 보조 입력·다음 달 계획 또는 납품처가 변경되었습니다. 화면을 다시 열어 검토한 뒤 저장하세요.")
            required = {row["id"] for row in state["destinations"] if row["active"] and row["required_for_report"]}
            if set(ids) != required:
                raise MonthlyInputValidationError("납품처 입력 행이 변경되었습니다. 화면을 다시 여세요.")
            _save(connection, current_values)
            _save(connection, next_values)
            for destination_id, q, c, item in validated:
                if q is None:
                    connection.execute("DELETE FROM monthly_plans WHERE report_month = ? AND destination_id = ?", (next_month, destination_id))
                else:
                    connection.execute(
                        "INSERT INTO monthly_plans(report_month, destination_id, quantity_ea_text, cost_won, representative_item) VALUES (?, ?, ?, ?, ?) "
                        "ON CONFLICT(report_month, destination_id) DO UPDATE SET quantity_ea_text = excluded.quantity_ea_text, cost_won = excluded.cost_won, representative_item = excluded.representative_item, updated_at = datetime('now')",
                        (next_month, destination_id, q, c, item),
                    )


def _values(inputs: ReportSupplementalInputs) -> tuple:
    _validate_month(inputs.report_month)
    if inputs.planned_sales_won is not None:
        _validate_money(inputs.planned_sales_won, "계획 매출")
    values = [inputs.report_month, inputs.planned_sales_won]
    for kind in ("planned", "actual"):
        quantity = getattr(inputs, f"nonregular_{kind}_quantity_ea")
        cost = getattr(inputs, f"nonregular_{kind}_cost_won")
        if (quantity is None) != (cost is None):
            raise MonthlyInputValidationError("비정규 수량과 비용을 함께 입력하세요.")
        if cost is not None:
            _validate_money(cost, "비정규 비용")
        values.extend((None if quantity is None else _canonical_quantity(quantity), cost))
    values.append(_optional_text(inputs.source_note, "입력 출처"))
    return tuple(values)


def _save(connection: sqlite3.Connection, values: tuple) -> None:
    connection.execute(
        "INSERT INTO report_supplemental_inputs(report_month, planned_sales_won, nonregular_planned_quantity_text, nonregular_planned_cost_won, nonregular_actual_quantity_text, nonregular_actual_cost_won, source_note) VALUES (?, ?, ?, ?, ?, ?, ?) "
        "ON CONFLICT(report_month) DO UPDATE SET planned_sales_won = excluded.planned_sales_won, nonregular_planned_quantity_text = excluded.nonregular_planned_quantity_text, nonregular_planned_cost_won = excluded.nonregular_planned_cost_won, nonregular_actual_quantity_text = excluded.nonregular_actual_quantity_text, nonregular_actual_cost_won = excluded.nonregular_actual_cost_won, source_note = excluded.source_note, updated_at = datetime('now')",
        values,
    )


def _get(connection: sqlite3.Connection, month: str) -> ReportSupplementalInputs | None:
    row = connection.execute("SELECT * FROM report_supplemental_inputs WHERE report_month = ?", (month,)).fetchone()
    if row is None:
        return None
    return ReportSupplementalInputs(
        row["report_month"], row["planned_sales_won"],
        None if row["nonregular_planned_quantity_text"] is None else Decimal(row["nonregular_planned_quantity_text"]),
        row["nonregular_planned_cost_won"],
        None if row["nonregular_actual_quantity_text"] is None else Decimal(row["nonregular_actual_quantity_text"]),
        row["nonregular_actual_cost_won"], row["source_note"], row["updated_at"],
    )


def _state(connection: sqlite3.Connection, month: str, next_month: str) -> dict:
    return {
        "report_month": month,
        "destinations": [dict(row) for row in connection.execute("SELECT * FROM destinations ORDER BY display_order, id")],
        "plans": [dict(row) for row in connection.execute("SELECT destination_id, quantity_ea_text, cost_won, representative_item FROM monthly_plans WHERE report_month = ? ORDER BY destination_id", (next_month,))],
        "supplemental": [dict(row) for row in connection.execute("SELECT report_month, planned_sales_won, nonregular_planned_quantity_text, nonregular_planned_cost_won, nonregular_actual_quantity_text, nonregular_actual_cost_won, source_note FROM report_supplemental_inputs WHERE report_month IN (?, ?) ORDER BY report_month", (month, next_month))],
    }


def _revision(state: dict) -> str:
    return hashlib.sha256(json.dumps(state, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()
