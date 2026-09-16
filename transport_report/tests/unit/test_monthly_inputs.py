from __future__ import annotations

import inspect
import sqlite3
from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError
from datetime import datetime, timedelta, timezone
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
        protocols.QuantityRecord("2026-09", first.id, Decimal("0"), "manual_erp")
    ]
    assert second.id not in [record.destination_id for record in provider.load("2026-09")]


def test_manual_provider_is_aggregate_duplicate_free_and_deterministic(
    repository, destinations, api
):
    protocols, monthly_inputs = api
    third, first, second = destinations
    repository.save_actual_quantity("2026-09", third.id, 1, "first note")
    repository.save_actual_quantity("2026-09", first.id, 2, "second note")
    repository.save_actual_quantity("2026-09", second.id, 3, None)
    repository.save_actual_quantity("2026-09", first.id, 4, "replacement note")

    records = monthly_inputs.ManualQuantityProvider(repository).load("2026-09")

    assert records == [
        protocols.QuantityRecord("2026-09", first.id, Decimal("4"), "manual_erp"),
        protocols.QuantityRecord("2026-09", second.id, Decimal("3"), "manual_erp"),
        protocols.QuantityRecord("2026-09", third.id, Decimal("1"), "manual_erp"),
    ]


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
        "2026-09-15",
    )
    updated = repository.save_sales(
        "2026-09",
        200_000,
        "   ",
        datetime(2026, 9, 16, 8, 5, 4, 999999, tzinfo=timezone.utc),
    )

    assert inserted.source_note == "finance team"
    assert inserted.confirmed_at == "2026-09-15T00:00:00+09:00"
    assert updated.amount_won == 200_000
    assert updated.source_note is None
    assert updated.confirmed_at == "2026-09-16T17:05:04+09:00"
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


def test_sqlite_integer_maximum_is_accepted_for_ids_money_and_revision(repository):
    maximum = 2**63 - 1
    with repository.database.connection() as connection:
        connection.execute(
            "INSERT INTO destinations(id, name, display_order) VALUES (?, ?, ?)",
            (maximum, "Maximum ID", 1),
        )
        connection.commit()

    quantity = repository.save_actual_quantity("2026-09", maximum, 1)
    sales = repository.save_sales("2026-09", maximum, None, "2026-09-15")
    run = repository.finalize_month("2026-09", maximum)

    assert quantity.destination_id == maximum
    assert sales.amount_won == maximum
    assert run.input_revision == maximum


def test_sqlite_integer_overflow_is_rejected_before_binding(repository, api):
    _, monthly_inputs = api
    overflow = 2**63
    operations = [
        lambda: repository.get_plan("2026-09", overflow),
        lambda: repository.save_sales("2026-09", overflow, None, "2026-09-15"),
        lambda: repository.finalize_month("2026-09", overflow),
    ]

    for operation in operations:
        with pytest.raises(monthly_inputs.MonthlyInputValidationError):
            operation()


@pytest.mark.parametrize(
    "confirmed_at",
    [None, "", "2026-09-15 12:30:00", "2026-09-15T12:30", 1],
)
def test_invalid_sales_confirmation_timestamp_is_rejected(repository, api, confirmed_at):
    _, monthly_inputs = api
    with pytest.raises(
        monthly_inputs.MonthlyInputValidationError, match="confirmed_at"
    ):
        repository.save_sales("2026-09", 1, "finance", confirmed_at)


def test_sales_timestamps_normalize_to_seoul_and_seconds(repository):
    from_string = repository.save_sales(
        "2026-09", 1, None, "2026-09-15T16:30:45.999999+00:00"
    )
    from_datetime = repository.save_sales(
        "2026-09",
        1,
        None,
        datetime(
            2026,
            9,
            15,
            12,
            30,
            45,
            123456,
            tzinfo=timezone(timedelta(hours=-4)),
        ),
    )

    assert from_string.confirmed_at == "2026-09-16T01:30:45+09:00"
    assert from_datetime.confirmed_at == from_string.confirmed_at


def test_sales_rejects_naive_datetime(repository, api):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="timezone"):
        repository.save_sales(
            "2026-09", 1, None, datetime(2026, 9, 15, 12, 30, 0)
        )


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


def test_clear_plan_and_sales_are_idempotent_and_distinguish_zero_from_missing(
    repository, destinations
):
    _, destination, _ = destinations
    repository.save_plan("2026-09", destination.id, 0, 0, None)
    repository.save_sales("2026-09", 0, None, "2026-09-15")
    assert repository.get_plan("2026-09", destination.id) is not None
    assert repository.get_sales("2026-09") is not None

    assert repository.clear_plan("2026-09", destination.id) is True
    assert repository.clear_plan("2026-09", destination.id) is False
    assert repository.clear_sales("2026-09") is True
    assert repository.clear_sales("2026-09") is False
    assert repository.get_plan("2026-09", destination.id) is None
    assert repository.get_sales("2026-09") is None


def test_locked_month_rejects_plan_and_sales_clears(repository, destinations, api):
    _, monthly_inputs = api
    _, destination, _ = destinations
    repository.save_plan("2026-09", destination.id, 0, 0, None)
    repository.save_sales("2026-09", 0, None, "2026-09-15")
    repository.finalize_month("2026-09", 1)

    with pytest.raises(monthly_inputs.MonthLockedError):
        repository.clear_plan("2026-09", destination.id)
    with pytest.raises(monthly_inputs.MonthLockedError):
        repository.clear_sales("2026-09")


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
    repository.finalize_month("2026-09", 1)

    operations = [
        lambda: repository.save_plan("2026-09", destination.id, 1, 1, None),
        lambda: repository.save_actual_quantity("2026-09", destination.id, 2),
        lambda: repository.clear_actual_quantity("2026-09", destination.id),
        lambda: repository.save_sales(
            "2026-09", 1, "finance", "2026-09-15"
        ),
    ]
    for operation in operations:
        with pytest.raises(monthly_inputs.MonthLockedError, match="2026-09"):
            operation()


def test_unlock_records_reason_and_time_then_allows_writes(repository, destinations):
    _, destination, _ = destinations
    locked = repository.finalize_month("2026-09", 1)
    unlocked = repository.unlock_month("2026-09", "  corrected finance input  ")

    assert locked.is_locked is True
    assert unlocked.id != locked.id
    assert unlocked.is_locked is False
    assert unlocked.locked_at is None
    assert unlocked.unlocked_at
    assert unlocked.unlock_reason == "corrected finance input"
    assert unlocked.input_revision == 1
    assert repository.is_month_locked("2026-09") is False

    repository.save_plan("2026-09", destination.id, 1, 1, None)
    repository.save_actual_quantity("2026-09", destination.id, 1)
    repository.save_sales("2026-09", 1, None, "2026-09-15")


def test_unlock_requires_a_locked_month_and_nonblank_reason(repository, api):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.MonthLockError, match="not locked"):
        repository.unlock_month("2026-09", "correction")

    repository.finalize_month("2026-09", 1)
    for reason in (None, "", "   "):
        with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="reason"):
            repository.unlock_month("2026-09", reason)
    assert repository.is_month_locked("2026-09") is True


@pytest.mark.parametrize("invalid_revision", [-1, True, 1.5, "1", " ", None])
def test_finalize_rejects_invalid_revision(repository, api, invalid_revision):
    _, monthly_inputs = api
    with pytest.raises(monthly_inputs.MonthlyInputValidationError, match="revision"):
        repository.finalize_month("2026-09", invalid_revision)


def test_finalize_rejects_already_locked_month(repository, api):
    _, monthly_inputs = api
    repository.finalize_month("2026-09", 1)
    with pytest.raises(monthly_inputs.MonthLockError, match="already locked"):
        repository.finalize_month("2026-09", 2)


def test_lock_state_is_single_row_and_lock_audit_is_append_only(repository):
    repository.finalize_month("2026-09", 1)
    repository.unlock_month("2026-09", "correction one")
    repository.finalize_month("2026-09", 2)
    repository.unlock_month("2026-09", "correction two")

    with repository.database.connection() as connection:
        locks = connection.execute(
            "SELECT * FROM month_locks WHERE report_month = ?", ("2026-09",)
        ).fetchall()
        events = connection.execute(
            "SELECT status, is_locked, input_revision, unlock_reason "
            "FROM report_runs WHERE report_month = ? ORDER BY id",
            ("2026-09",),
        ).fetchall()

    assert len(locks) == 1
    assert locks[0]["is_locked"] == 0
    assert locks[0]["input_revision"] == 2
    assert locks[0]["last_unlock_reason"] == "correction two"
    assert [row["status"] for row in events] == [
        "month_locked",
        "month_unlocked",
        "month_locked",
        "month_unlocked",
    ]


def test_concurrent_lock_transitions_have_one_winner(repository, api):
    _, monthly_inputs = api

    with ThreadPoolExecutor(max_workers=2) as executor:
        finalize_results = [
            future.exception()
            for future in (
                executor.submit(repository.finalize_month, "2026-09", 1),
                executor.submit(repository.finalize_month, "2026-09", 2),
            )
        ]
    assert sum(result is None for result in finalize_results) == 1
    assert sum(isinstance(result, monthly_inputs.MonthLockError) for result in finalize_results) == 1

    with ThreadPoolExecutor(max_workers=2) as executor:
        unlock_results = [
            future.exception()
            for future in (
                executor.submit(repository.unlock_month, "2026-09", "first"),
                executor.submit(repository.unlock_month, "2026-09", "second"),
            )
        ]
    assert sum(result is None for result in unlock_results) == 1
    assert sum(isinstance(result, monthly_inputs.MonthLockError) for result in unlock_results) == 1

    with repository.database.connection() as connection:
        statuses = [
            row["status"]
            for row in connection.execute(
                "SELECT status FROM report_runs WHERE report_month = ? ORDER BY id",
                ("2026-09",),
            )
        ]
    assert statuses == ["month_locked", "month_unlocked"]


def test_unlock_does_not_rewrite_a_future_report_generation_row(repository):
    repository.finalize_month("2026-09", 1)
    with repository.database.connection() as connection:
        future_run_id = connection.execute(
            "INSERT INTO report_runs"
            "(report_month, status, is_locked, locked_at, input_revision, "
            "output_hashes_json) VALUES (?, ?, 1, ?, ?, ?)",
            (
                "2026-09",
                "report_generated",
                "2026-09-20T10:00:00+09:00",
                "1",
                '{"deck":"sha256"}',
            ),
        ).lastrowid
        connection.commit()

    repository.unlock_month("2026-09", "correct inputs")

    with repository.database.connection() as connection:
        future_run = connection.execute(
            "SELECT * FROM report_runs WHERE id = ?", (future_run_id,)
        ).fetchone()
    assert future_run["status"] == "report_generated"
    assert future_run["is_locked"] == 1
    assert future_run["unlocked_at"] is None
    assert future_run["unlock_reason"] is None
    assert future_run["output_hashes_json"] == '{"deck":"sha256"}'


def test_month_lock_audit_rows_are_immutable_but_other_report_runs_are_not(repository):
    locked = repository.finalize_month("2026-09", 1)
    repository.unlock_month("2026-09", "correct inputs")

    with repository.database.connection() as connection:
        with pytest.raises(sqlite3.IntegrityError, match="month lock audit immutable"):
            connection.execute(
                "UPDATE report_runs SET status = 'changed' WHERE id = ?", (locked.id,)
            )
        with pytest.raises(sqlite3.IntegrityError, match="month lock audit immutable"):
            connection.execute("DELETE FROM report_runs WHERE id = ?", (locked.id,))
        generated_id = connection.execute(
            "INSERT INTO report_runs(report_month, status) VALUES (?, ?)",
            ("2026-09", "generating"),
        ).lastrowid
        connection.execute(
            "UPDATE report_runs SET status = 'generated' WHERE id = ?",
            (generated_id,),
        )
        assert connection.execute(
            "SELECT status FROM report_runs WHERE id = ?", (generated_id,)
        ).fetchone()[0] == "generated"
        with pytest.raises(sqlite3.IntegrityError, match="month lock audit immutable"):
            connection.execute(
                "UPDATE report_runs SET status = 'month_locked' WHERE id = ?",
                (generated_id,),
            )


def test_database_triggers_block_non_repository_month_writes_until_unlock(
    repository, destinations
):
    _, destination, second_destination = destinations
    with repository.database.connection() as connection:
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-09", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid
        cost_id = connection.execute(
            "INSERT INTO monthly_actual_costs"
            "(report_month, destination_id, cost_won, source_type) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09", destination.id, 1000, "manual"),
        ).lastrowid
        entry_id = connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-09",
                destination.id,
                "Sheet1",
                2,
                1,
                "regular",
                "1",
                1000,
            ),
        ).lastrowid
        connection.commit()
    repository.finalize_month("2026-09", 1)

    with repository.database.connection() as connection:
        locked_writes = [
            (
                "INSERT INTO monthly_actual_costs"
                "(report_month, destination_id, cost_won, source_type) "
                "VALUES (?, ?, ?, ?)",
                ("2026-09", second_destination.id, 1000, "manual"),
            ),
            ("UPDATE monthly_actual_costs SET cost_won = 2000 WHERE id = ?", (cost_id,)),
            ("DELETE FROM monthly_actual_costs WHERE id = ?", (cost_id,)),
            (
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_sheet, "
                "source_row, transport_day, transport_type, trip_count_text, cost_won) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    batch_id,
                    "2026-09",
                    destination.id,
                    "Sheet1",
                    3,
                    1,
                    "regular",
                    "1",
                    1000,
                ),
            ),
            ("UPDATE transport_entries SET cost_won = 2000 WHERE id = ?", (entry_id,)),
            ("DELETE FROM transport_entries WHERE id = ?", (entry_id,)),
        ]
        for sql, parameters in locked_writes:
            with pytest.raises(sqlite3.IntegrityError, match="report month locked"):
                connection.execute(sql, parameters)

    repository.unlock_month("2026-09", "correct inputs")
    with repository.database.connection() as connection:
        connection.execute(
            "INSERT INTO monthly_actual_costs"
            "(report_month, destination_id, cost_won, source_type) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09", second_destination.id, 1000, "manual"),
        )
        connection.execute(
            "UPDATE monthly_actual_costs SET cost_won = 2000 WHERE id = ?", (cost_id,)
        )
        connection.execute("DELETE FROM monthly_actual_costs WHERE id = ?", (cost_id,))
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-09",
                destination.id,
                "Sheet1",
                3,
                1,
                "regular",
                "1",
                1000,
            ),
        )
        connection.execute(
            "UPDATE transport_entries SET cost_won = 2000 WHERE id = ?", (entry_id,)
        )
        connection.execute("DELETE FROM transport_entries WHERE id = ?", (entry_id,))
        connection.commit()


def test_database_triggers_cover_plan_quantity_sales_and_leave_batches_writable(
    repository, destinations
):
    _, destination, _ = destinations
    with repository.database.connection() as connection:
        connection.execute(
            "INSERT INTO monthly_plans"
            "(report_month, destination_id, quantity_ea_text, cost_won) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09", destination.id, "1", 1000),
        )
        connection.execute(
            "INSERT INTO monthly_actual_quantities"
            "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
            ("2026-09", destination.id, "1"),
        )
        connection.execute(
            "INSERT INTO monthly_sales"
            "(report_month, amount_won, confirmed_at) VALUES (?, ?, ?)",
            ("2026-09", 1000, "2026-09-15T00:00:00+09:00"),
        )
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-09", "test", "source.xlsx", "b" * 64, "staged"),
        ).lastrowid
        connection.commit()
    repository.finalize_month("2026-09", 1)

    locked_operations = [
        (
            "INSERT INTO monthly_plans"
            "(report_month, destination_id, quantity_ea_text, cost_won) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09", destination.id, "2", 2000),
        ),
        (
            "UPDATE monthly_plans SET cost_won = 2000 "
            "WHERE report_month = ? AND destination_id = ?",
            ("2026-09", destination.id),
        ),
        (
            "DELETE FROM monthly_plans WHERE report_month = ? AND destination_id = ?",
            ("2026-09", destination.id),
        ),
        (
            "INSERT INTO monthly_actual_quantities"
            "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
            ("2026-09", destination.id, "2"),
        ),
        (
            "UPDATE monthly_actual_quantities SET quantity_ea_text = '2' "
            "WHERE report_month = ? AND destination_id = ?",
            ("2026-09", destination.id),
        ),
        (
            "DELETE FROM monthly_actual_quantities "
            "WHERE report_month = ? AND destination_id = ?",
            ("2026-09", destination.id),
        ),
        (
            "INSERT INTO monthly_sales(report_month, amount_won) VALUES (?, ?)",
            ("2026-09", 2000),
        ),
        (
            "UPDATE monthly_sales SET amount_won = 2000 WHERE report_month = ?",
            ("2026-09",),
        ),
        ("DELETE FROM monthly_sales WHERE report_month = ?", ("2026-09",)),
    ]
    with repository.database.connection() as connection:
        for sql, parameters in locked_operations:
            with pytest.raises(sqlite3.IntegrityError, match="report month locked"):
                connection.execute(sql, parameters)
        connection.execute(
            "UPDATE import_batches SET status = 'validated' WHERE id = ?", (batch_id,)
        )
        connection.commit()

    repository.unlock_month("2026-09", "correct inputs")
    with repository.database.connection() as connection:
        for table in (
            "monthly_plans",
            "monthly_actual_quantities",
            "monthly_sales",
        ):
            connection.execute(
                f"DELETE FROM {table} WHERE report_month = ?", ("2026-09",)
            )
        connection.execute(
            "INSERT INTO monthly_plans"
            "(report_month, destination_id, quantity_ea_text, cost_won) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09", destination.id, "2", 2000),
        )
        connection.execute(
            "INSERT INTO monthly_actual_quantities"
            "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
            ("2026-09", destination.id, "2"),
        )
        connection.execute(
            "INSERT INTO monthly_sales(report_month, amount_won) VALUES (?, ?)",
            ("2026-09", 2000),
        )
        connection.commit()


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
