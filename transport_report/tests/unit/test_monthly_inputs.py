from __future__ import annotations

import inspect
import sqlite3
from dataclasses import FrozenInstanceError
from datetime import datetime, timezone
from decimal import Decimal
from importlib import import_module

import pytest

from app.db import Database
from app.repositories.masters import MasterRepository


@pytest.fixture
def api():
    try:
        protocols = import_module("app.importers.protocols")
        monthly_inputs = import_module("app.repositories.monthly_inputs")
    except ModuleNotFoundError as error:
        pytest.fail(f"monthly input API is not implemented: {error}")
    return protocols, monthly_inputs


@pytest.fixture
def repository(tmp_path, api):
    _, monthly_inputs = api
    database = Database(tmp_path / "app.db")
    database.migrate()
    return monthly_inputs.MonthlyInputRepository(database)


@pytest.fixture
def destinations(repository):
    masters = MasterRepository(repository.database)
    return (
        masters.create_destination("Third", 30),
        masters.create_destination("First", 10),
        masters.create_destination("Second", 20),
    )


def test_quantity_protocol_supports_runtime_structural_use(api):
    protocols, _ = api

    class StubProvider:
        def load(self, report_month: str):
            return []

    assert isinstance(StubProvider(), protocols.QuantityProvider)
    record = protocols.QuantityRecord("2026-09", 1, Decimal("1.5"), "manual")
    assert record.quantity_ea == Decimal("1.5")
    with pytest.raises(FrozenInstanceError):
        record.quantity_ea = Decimal("2")


def test_manual_provider_preserves_explicit_zero_and_omits_absent_destinations(
    repository, destinations, api
):
    protocols, monthly_inputs = api
    _, first, second = destinations
    repository.save_actual_quantity("2026-09", first.id, Decimal("0"), " checked ")

    provider = monthly_inputs.ManualQuantityProvider(repository)

    assert isinstance(provider, protocols.QuantityProvider)
    assert provider.load("2026-09") == [
        protocols.QuantityRecord("2026-09", first.id, Decimal("0"), "checked")
    ]
    assert second.id not in [record.destination_id for record in provider.load("2026-09")]


def test_plan_insert_update_round_trip_and_trimmed_optional_item(repository, destinations):
    _, destination, _ = destinations

    inserted = repository.save_plan(
        "2026-09", destination.id, Decimal("1.250"), 1000, "  ITEM-A  "
    )
    updated = repository.save_plan(
        "2026-09", destination.id, Decimal("2"), 2000, "   "
    )

    assert inserted.quantity_ea == Decimal("1.25")
    assert inserted.representative_item == "ITEM-A"
    assert updated.quantity_ea == Decimal("2")
    assert updated.cost_won == 2000
    assert updated.representative_item is None
    assert repository.get_plan("2026-09", destination.id) == updated


def test_actual_quantity_insert_update_round_trip(repository, destinations):
    _, destination, _ = destinations

    inserted = repository.save_actual_quantity(
        "2026-09", destination.id, Decimal("1.250"), "  first check  "
    )
    updated = repository.save_actual_quantity("2026-09", destination.id, 0)

    assert inserted.quantity_ea == Decimal("1.25")
    assert inserted.source_note == "first check"
    assert updated.quantity_ea == Decimal("0")
    assert updated.source_note == "ERP 수기 확인"
    assert repository.get_actual_quantity("2026-09", destination.id) == updated


def test_sales_insert_update_round_trip(repository):
    inserted = repository.save_sales(
        "2026-09",
        100_000,
        "  finance team  ",
        "2026-09-15T12:30:00",
    )
    updated = repository.save_sales(
        "2026-09",
        200_000,
        "   ",
        datetime(2026, 9, 16, 8, 5, 4, 999999, tzinfo=timezone.utc),
    )

    assert inserted.source_note == "finance team"
    assert updated.amount_won == 200_000
    assert updated.source_note is None
    assert updated.confirmed_at == "2026-09-16T08:05:04+00:00"
    assert repository.get_sales("2026-09") == updated


def test_decimal_values_are_stored_in_canonical_text(repository, destinations):
    _, destination, _ = destinations
    repository.save_plan("2026-09", destination.id, Decimal("0.00"), 0, None)
    repository.save_actual_quantity("2026-09", destination.id, Decimal("1.250"))

    with repository.database.connection() as connection:
        plan_text = connection.execute(
            "SELECT quantity_ea_text FROM monthly_plans"
        ).fetchone()[0]
        actual_text = connection.execute(
            "SELECT quantity_ea_text FROM monthly_actual_quantities"
        ).fetchone()[0]

    assert plan_text == "0"
    assert actual_text == "1.25"


@pytest.mark.parametrize(
    "month",
    ["2026-00", "2026-13", "2026-1", "２０２６-09", "2026-09 ", 202609, None],
)
def test_invalid_report_month_is_rejected_without_writes(repository, destinations, api, month):
    _, monthly_inputs = api
    _, destination, _ = destinations

    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="YYYY-MM"):
        repository.save_plan(month, destination.id, Decimal("1"), 1, None)

    assert repository.list_plans("2026-09") == []


@pytest.mark.parametrize("destination_id", [0, -1, True, 1.5, "1", None])
def test_invalid_destination_id_is_rejected(repository, api, destination_id):
    _, monthly_inputs = api
    with pytest.raises(
        monthly_inputs.MonthlyInputValidationError, match="positive integer"
    ):
        repository.save_actual_quantity("2026-09", destination_id, Decimal("1"))


@pytest.mark.parametrize(
    "quantity",
    [Decimal("-1"), Decimal("NaN"), Decimal("Infinity"), True, 1.5, "1", None],
)
def test_invalid_quantities_are_rejected(repository, destinations, api, quantity):
    _, monthly_inputs = api
    _, destination, _ = destinations
    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="quantity"):
        repository.save_actual_quantity("2026-09", destination.id, quantity)


@pytest.mark.parametrize("money", [-1, True, 1.5, "1", Decimal("1"), None])
def test_invalid_money_is_rejected(repository, destinations, api, money):
    _, monthly_inputs = api
    _, destination, _ = destinations
    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="won"):
        repository.save_plan("2026-09", destination.id, Decimal("1"), money, None)


@pytest.mark.parametrize(
    "confirmed_at",
    [None, "", "2026-09-15", "2026-09-15 12:30:00", "2026-09-15T12:30", 1],
)
def test_invalid_sales_confirmation_timestamp_is_rejected(repository, api, confirmed_at):
    _, monthly_inputs = api
    with pytest.raises(
        monthly_inputs.MonthlyInputValidationError, match="confirmed_at"
    ):
        repository.save_sales("2026-09", 1, "finance", confirmed_at)


def test_invalid_text_values_are_rejected(repository, destinations, api):
    _, monthly_inputs = api
    _, destination, _ = destinations
    operations = [
        lambda: repository.save_plan(
            "2026-09", destination.id, Decimal("1"), 1, 123
        ),
        lambda: repository.save_actual_quantity(
            "2026-09", destination.id, Decimal("1"), 123
        ),
        lambda: repository.save_sales(
            "2026-09", 1, 123, "2026-09-15T12:30:00"
        ),
    ]
    for operation in operations:
        with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="text"):
            operation()


def test_missing_destination_is_translated(repository, api):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.DestinationNotFoundError, match="999999"):
        repository.save_plan("2026-09", 999999, Decimal("1"), 1, None)
    with pytest.raises(monthly_inputs.DestinationNotFoundError, match="999999"):
        repository.save_actual_quantity("2026-09", 999999, Decimal("1"))


def test_database_fk_error_is_translated(repository, api, monkeypatch):
    _, monthly_inputs = api
    monkeypatch.setattr(repository, "_require_destination", lambda *_: None)

    with pytest.raises(monthly_inputs.DestinationNotFoundError, match="999999"):
        repository.save_plan("2026-09", 999999, Decimal("1"), 1, None)


def test_clear_actual_quantity_distinguishes_missing_from_zero(repository, destinations):
    _, destination, _ = destinations
    repository.save_actual_quantity("2026-09", destination.id, Decimal("0"))
    assert repository.get_actual_quantity("2026-09", destination.id) is not None

    assert repository.clear_actual_quantity("2026-09", destination.id) is True
    assert repository.get_actual_quantity("2026-09", destination.id) is None
    assert repository.clear_actual_quantity("2026-09", destination.id) is False


def test_month_lists_use_destination_display_order(repository, destinations):
    third, first, second = destinations
    for destination in (third, first, second):
        repository.save_plan("2026-09", destination.id, 1, 1, None)
        repository.save_actual_quantity("2026-09", destination.id, 1)

    assert [record.destination_id for record in repository.list_plans("2026-09")] == [
        first.id,
        second.id,
        third.id,
    ]
    assert [
        record.destination_id
        for record in repository.list_actual_quantities("2026-09")
    ] == [first.id, second.id, third.id]


def test_finalized_month_rejects_every_ordinary_write_and_clear(
    repository, destinations, api
):
    _, monthly_inputs = api
    _, destination, _ = destinations
    repository.save_actual_quantity("2026-09", destination.id, 1)
    repository.finalize_month("2026-09", "input-revision-1")

    operations = [
        lambda: repository.save_plan("2026-09", destination.id, 1, 1, None),
        lambda: repository.save_actual_quantity("2026-09", destination.id, 2),
        lambda: repository.clear_actual_quantity("2026-09", destination.id),
        lambda: repository.save_sales(
            "2026-09", 1, "finance", "2026-09-15T12:30:00"
        ),
    ]
    for operation in operations:
        with pytest.raises(monthly_inputs.MonthLockedError, match="2026-09"):
            operation()


def test_unlock_records_reason_and_time_then_allows_writes(repository, destinations):
    _, destination, _ = destinations
    locked = repository.finalize_month("2026-09", "  revision-1  ")
    unlocked = repository.unlock_month("2026-09", "  corrected finance input  ")

    assert locked.is_locked is True
    assert unlocked.id == locked.id
    assert unlocked.is_locked is False
    assert unlocked.locked_at
    assert unlocked.unlocked_at
    assert unlocked.unlock_reason == "corrected finance input"
    assert unlocked.input_revision == "revision-1"
    assert repository.is_month_locked("2026-09") is False

    repository.save_plan("2026-09", destination.id, 1, 1, None)
    repository.save_actual_quantity("2026-09", destination.id, 1)
    repository.save_sales("2026-09", 1, None, "2026-09-15T12:30:00")


def test_unlock_requires_a_locked_month_and_nonblank_reason(repository, api):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.MonthLockError, match="not locked"):
        repository.unlock_month("2026-09", "correction")

    repository.finalize_month("2026-09", "revision-1")
    for reason in (None, "", "   "):
        with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="reason"):
            repository.unlock_month("2026-09", reason)
    assert repository.is_month_locked("2026-09") is True


def test_finalize_rejects_blank_revision_and_already_locked_month(repository, api):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="revision"):
        repository.finalize_month("2026-09", " ")
    repository.finalize_month("2026-09", "revision-1")
    with pytest.raises(monthly_inputs.MonthLockError, match="already locked"):
        repository.finalize_month("2026-09", "revision-2")


class TrackingConnection(sqlite3.Connection):
    rollback_calls = 0
    close_calls = 0
    fail_commit = False

    def commit(self):
        if self.fail_commit:
            raise sqlite3.OperationalError("forced commit failure")
        return super().commit()

    def rollback(self):
        type(self).rollback_calls += 1
        return super().rollback()

    def close(self):
        type(self).close_calls += 1
        return super().close()


class TrackingDatabase(Database):
    def connect(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.path, factory=TrackingConnection)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection


def test_write_failure_rolls_back_and_closes_connection(tmp_path, api):
    _, monthly_inputs = api
    database = TrackingDatabase(tmp_path / "app.db")
    database.migrate()
    destination = MasterRepository(database).create_destination("Destination", 1)
    TrackingConnection.rollback_calls = 0
    TrackingConnection.close_calls = 0
    TrackingConnection.fail_commit = True

    try:
        with pytest.raises(sqlite3.OperationalError, match="forced commit"):
            monthly_inputs.MonthlyInputRepository(database).save_plan(
                "2026-09", destination.id, 1, 1, None
            )
    finally:
        TrackingConnection.fail_commit = False

    assert TrackingConnection.rollback_calls == 1
    assert TrackingConnection.close_calls == 1
    with database.connection() as connection:
        assert connection.execute("SELECT COUNT(*) FROM monthly_plans").fetchone()[0] == 0


def test_sales_is_month_level_only(repository):
    assert "destination_id" not in inspect.signature(repository.save_sales).parameters
    with repository.database.connection() as connection:
        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(monthly_sales)").fetchall()
        }
    assert "destination_id" not in columns
