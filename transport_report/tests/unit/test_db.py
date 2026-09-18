import shutil
import sqlite3
import threading
from concurrent.futures import ThreadPoolExecutor

import pytest

import app.db as db_module
from app.db import Database
from app.repositories.monthly_inputs import MonthlyInputRepository


def test_initial_migration_creates_required_tables(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    names = db.table_names()
    assert {
        "destinations",
        "destination_aliases",
        "vehicle_rates",
        "report_groups",
        "report_group_members",
        "monthly_plans",
        "monthly_actual_quantities",
        "monthly_actual_costs",
        "monthly_sales",
        "transport_entries",
        "import_batches",
        "report_runs",
        "month_locks",
    } <= names


def test_import_batch_selection_migration_adds_strict_current_invariant(tmp_path):
    database = Database(tmp_path / "app.db")
    database.migrate()

    with database.connection() as connection:
        columns = {
            row["name"]: row
            for row in connection.execute("PRAGMA table_info(import_batches)")
        }
        indexes = {
            row["name"]
            for row in connection.execute("PRAGMA index_list(import_batches)")
        }
        first_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status, "
            "is_current) VALUES (?, ?, ?, ?, ?, ?)",
            ("2026-08", "transport", "a.xlsx", "a" * 64, "imported", 1),
        ).lastrowid
        connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-09", "transport", "b.xlsx", "b" * 64, "imported"),
        )

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO import_batches"
                "(report_month, source_type, source_filename, file_sha256, status, "
                "is_current) VALUES (?, ?, ?, ?, ?, ?)",
                ("2026-08", "transport", "c.xlsx", "c" * 64, "imported", 1),
            )
        for invalid in (2, -1, "true"):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    "UPDATE import_batches SET is_current = ? WHERE id = ?",
                    (invalid, first_id),
                )

        default_state = connection.execute(
            "SELECT is_current FROM import_batches WHERE report_month = '2026-09'"
        ).fetchone()[0]

    assert columns["is_current"]["notnull"] == 1
    assert columns["is_current"]["dflt_value"] == "0"
    assert "uq_import_batches_current_month_source" in indexes
    assert default_state == 0


def test_import_batch_selection_upgrade_backfills_latest_eligible_and_preserves_data(
    tmp_path,
):
    bundled = Database(tmp_path / "unused.db").migrations_dir
    first_four = tmp_path / "first-four"
    first_four.mkdir()
    for filename in (
        "001_initial.sql",
        "002_month_locks.sql",
        "003_transport_entry_source_alias.sql",
        "004_transport_entry_source_date.sql",
    ):
        shutil.copyfile(bundled / filename, first_four / filename)

    database_path = tmp_path / "app.db"
    legacy = Database(database_path, migrations_dir=first_four)
    legacy.migrate()
    with legacy.connection() as connection:
        batch_values = (
            ("2026-08", "transport", "old.xlsx", "a" * 64, "imported", None),
            (
                "2026-08",
                "transport",
                "corrected.xlsx",
                "b" * 64,
                "imported_with_errors",
                "Unknown destination alias: 미등록",
            ),
            ("2026-08", "transport", "staged.xlsx", "c" * 64, "staged", None),
            ("2026-08", "erp", "failed.xlsx", "d" * 64, "failed", "failed"),
            ("2026-09", "transport", "next.xlsx", "e" * 64, "imported", None),
        )
        batch_ids = []
        for values in batch_values:
            batch_ids.append(
                connection.execute(
                    "INSERT INTO import_batches"
                    "(report_month, source_type, source_filename, file_sha256, "
                    "status, error_summary) VALUES (?, ?, ?, ?, ?, ?)",
                    values,
                ).lastrowid
            )
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, unresolved_alias, source_alias, "
            "source_sheet, source_row, source_date, transport_day, "
            "transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_ids[1],
                "2026-08",
                "미등록",
                "미등록",
                "정규운송",
                7,
                "2026-08-01",
                1,
                "regular",
                "1",
                1000,
            ),
        )
        connection.commit()

    upgraded = Database(database_path)
    upgraded.migrate()

    with upgraded.connection() as connection:
        batches = [
            tuple(row)
            for row in connection.execute(
                "SELECT id, status, error_summary, is_current "
                "FROM import_batches ORDER BY id"
            )
        ]
        entry = connection.execute(
            "SELECT import_batch_id, unresolved_alias, cost_won "
            "FROM transport_entries"
        ).fetchone()
        versions = [
            row["version"]
            for row in connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            )
        ]

    assert [row[3] for row in batches] == [0, 1, 0, 0, 1]
    assert batches[1][2] == "Unknown destination alias: 미등록"
    assert tuple(entry) == (batch_ids[1], "미등록", 1000)
    assert versions == [1, 2, 3, 4, 5]


def test_failed_import_batch_selection_migration_rolls_back_column_and_version(
    tmp_path,
):
    bundled = Database(tmp_path / "unused.db").migrations_dir
    migrations = tmp_path / "failing-migrations"
    migrations.mkdir()
    for filename in (
        "001_initial.sql",
        "002_month_locks.sql",
        "003_transport_entry_source_alias.sql",
        "004_transport_entry_source_date.sql",
    ):
        shutil.copyfile(bundled / filename, migrations / filename)
    migration_005 = (bundled / "005_import_batch_selection.sql").read_text(
        encoding="utf-8"
    )
    (migrations / "005_import_batch_selection.sql").write_text(
        migration_005
        + "\nINSERT INTO table_that_does_not_exist(value) VALUES (1);\n",
        encoding="utf-8",
    )
    database = Database(tmp_path / "app.db", migrations_dir=migrations)

    with pytest.raises(sqlite3.OperationalError, match="no such table"):
        database.migrate()

    with database.connection() as connection:
        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(import_batches)")
        }
        versions = [
            row["version"]
            for row in connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            )
        ]

    assert "is_current" not in columns
    assert versions == [1, 2, 3, 4]


def test_month_lock_migration_is_versioned_and_upgrades_an_existing_database(tmp_path):
    migrations_dir = Database(tmp_path / "unused.db").migrations_dir
    first_only = tmp_path / "first-only"
    first_only.mkdir()
    shutil.copyfile(
        migrations_dir / "001_initial.sql", first_only / "001_initial.sql"
    )
    database_path = tmp_path / "app.db"
    Database(database_path, migrations_dir=first_only).migrate()

    database = Database(database_path)
    database.migrate()

    with database.connection() as connection:
        versions = [
            row["version"]
            for row in connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            )
        ]
        lock_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(month_locks)")
        }
        trigger_names = {
            row["name"]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'trigger'"
            )
        }
    assert versions == [1, 2, 3, 4, 5]
    assert {
        "report_month",
        "is_locked",
        "locked_at",
        "unlocked_at",
        "input_revision",
        "last_unlock_reason",
    } <= lock_columns
    assert {
        f"{table}_reject_locked_{operation}"
        for table in (
            "monthly_plans",
            "monthly_actual_quantities",
            "monthly_sales",
            "monthly_actual_costs",
            "transport_entries",
        )
        for operation in ("insert", "update", "delete")
    } <= trigger_names


def test_transport_source_alias_migration_preserves_known_legacy_state(tmp_path):
    migrations_dir = Database(tmp_path / "unused.db").migrations_dir
    first_two = tmp_path / "first-two"
    first_two.mkdir()
    for filename in ("001_initial.sql", "002_month_locks.sql"):
        shutil.copyfile(migrations_dir / filename, first_two / filename)

    database_path = tmp_path / "app.db"
    legacy = Database(database_path, migrations_dir=first_two)
    legacy.migrate()
    with legacy.connection() as connection:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Legacy Destination", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "transport", "legacy.xlsx", "a" * 64, "imported"),
        ).lastrowid
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-08",
                destination_id,
                "Sheet1",
                2,
                1,
                "regular",
                "1",
                1_000,
            ),
        )
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, unresolved_alias, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-08",
                "Legacy Unknown",
                "Sheet1",
                3,
                2,
                "regular",
                "1",
                2_000,
            ),
        )
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, unresolved_alias, "
            "source_sheet, source_row, transport_day, transport_type, "
            "trip_count_text, cost_won) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-08",
                destination_id,
                "   ",
                "Sheet1",
                4,
                3,
                "regular",
                "1",
                3_000,
            ),
        )
        connection.execute(
            "INSERT INTO month_locks"
            "(report_month, is_locked, locked_at, input_revision) "
            "VALUES (?, 1, ?, ?)",
            ("2026-08", "2026-09-01T00:00:00", "legacy-revision"),
        )
        locked_update_trigger_before = connection.execute(
            "SELECT sql FROM sqlite_master "
            "WHERE type = 'trigger' "
            "AND name = 'transport_entries_reject_locked_update'"
        ).fetchone()["sql"]
        connection.commit()

    upgraded = Database(database_path)
    upgraded.migrate()

    with upgraded.connection() as connection:
        rows = connection.execute(
            "SELECT destination_id, unresolved_alias, source_alias, source_date "
            "FROM transport_entries ORDER BY source_row"
        ).fetchall()
        versions = [
            row["version"]
            for row in connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            )
        ]
        lock = connection.execute(
            "SELECT is_locked, input_revision FROM month_locks "
            "WHERE report_month = '2026-08'"
        ).fetchone()
        locked_update_trigger_after = connection.execute(
            "SELECT sql FROM sqlite_master "
            "WHERE type = 'trigger' "
            "AND name = 'transport_entries_reject_locked_update'"
        ).fetchone()["sql"]
        with pytest.raises(sqlite3.IntegrityError, match="report month locked"):
            connection.execute(
                "UPDATE transport_entries SET cost_won = cost_won + 1 "
                "WHERE source_row = 3"
            )
    assert versions == [1, 2, 3, 4, 5]
    assert tuple(rows[0]) == (destination_id, None, None, None)
    assert tuple(rows[1]) == (None, "Legacy Unknown", "Legacy Unknown", None)
    assert tuple(rows[2]) == (destination_id, "   ", None, None)
    assert tuple(lock) == (1, "legacy-revision")
    assert locked_update_trigger_after == locked_update_trigger_before


def test_failed_migration_rolls_back_dropped_locked_update_trigger(tmp_path):
    bundled = Database(tmp_path / "unused.db").migrations_dir
    migrations = tmp_path / "failing-migrations"
    migrations.mkdir()
    for filename in ("001_initial.sql", "002_month_locks.sql"):
        shutil.copyfile(bundled / filename, migrations / filename)
    database = Database(tmp_path / "app.db", migrations_dir=migrations)
    database.migrate()
    with database.connection() as connection:
        trigger_before = connection.execute(
            "SELECT sql FROM sqlite_master "
            "WHERE name = 'transport_entries_reject_locked_update'"
        ).fetchone()["sql"]

    (migrations / "003_fails.sql").write_text(
        "DROP TRIGGER transport_entries_reject_locked_update;\n"
        "INSERT INTO table_that_does_not_exist(value) VALUES (1);\n",
        encoding="utf-8",
    )

    with pytest.raises(sqlite3.OperationalError, match="no such table"):
        database.migrate()

    with database.connection() as connection:
        trigger_after = connection.execute(
            "SELECT sql FROM sqlite_master "
            "WHERE name = 'transport_entries_reject_locked_update'"
        ).fetchone()["sql"]
        versions = [
            row["version"]
            for row in connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            )
        ]
    assert trigger_after == trigger_before
    assert versions == [1, 2]


def test_transport_source_alias_rejects_blank_nonnull_values(tmp_path):
    database = Database(tmp_path / "app.db")
    database.migrate()
    with database.connection() as connection:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "transport", "source.xlsx", "b" * 64, "imported"),
        ).lastrowid

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_alias, "
                "source_sheet, source_row, transport_day, transport_type, "
                "trip_count_text, cost_won) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    batch_id,
                    "2026-08",
                    destination_id,
                    "   ",
                    "Sheet1",
                    2,
                    1,
                    "regular",
                    "1",
                    1_000,
                ),
            )


@pytest.mark.parametrize(
    "revision", ["revision-1", "sha256:abc123", "  revision-1  "]
)
def test_month_lock_migration_backfills_opaque_revision_and_appends_audit(
    tmp_path, revision
):
    migrations_dir = Database(tmp_path / "unused.db").migrations_dir
    first_only = tmp_path / "first-only"
    first_only.mkdir()
    shutil.copyfile(
        migrations_dir / "001_initial.sql", first_only / "001_initial.sql"
    )
    database_path = tmp_path / "app.db"
    legacy_database = Database(database_path, migrations_dir=first_only)
    legacy_database.migrate()
    with legacy_database.connection() as connection:
        original_id = connection.execute(
            "INSERT INTO report_runs"
            "(report_month, status, is_locked, locked_at, input_revision) "
            "VALUES (?, ?, 1, ?, ?)",
            ("2026-09", "finalized", "2026-09-20T10:00:00+09:00", revision),
        ).lastrowid
        connection.commit()

    upgraded = Database(database_path)
    upgraded.migrate()

    with upgraded.connection() as connection:
        lock_row = connection.execute(
            "SELECT * FROM month_locks WHERE report_month = ?", ("2026-09",)
        ).fetchone()
        original = connection.execute(
            "SELECT * FROM report_runs WHERE id = ?", (original_id,)
        ).fetchone()
        events = connection.execute(
            "SELECT * FROM report_runs WHERE id <> ? ORDER BY id", (original_id,)
        ).fetchall()
    assert lock_row["is_locked"] == 1
    assert lock_row["locked_at"] == "2026-09-20T10:00:00+09:00"
    assert lock_row["input_revision"] == revision
    assert original["status"] == "finalized"
    assert original["input_revision"] == revision
    assert [row["status"] for row in events] == ["month_locked"]
    assert events[0]["input_revision"] == revision

    MonthlyInputRepository(upgraded).unlock_month("2026-09", "correct legacy input")
    with upgraded.connection() as connection:
        audit_rows = connection.execute(
            "SELECT id, status, input_revision FROM report_runs "
            "WHERE status IN ('month_locked', 'month_unlocked') ORDER BY id"
        ).fetchall()
        for audit_row in audit_rows:
            for statement in (
                "UPDATE report_runs SET status = 'changed' WHERE id = ?",
                "DELETE FROM report_runs WHERE id = ?",
                "INSERT OR REPLACE INTO report_runs"
                "(id, report_month, status) VALUES (?, '2026-09', 'changed')",
            ):
                with pytest.raises(sqlite3.IntegrityError, match="immutable"):
                    connection.execute(statement, (audit_row["id"],))
    assert [row["status"] for row in audit_rows] == [
        "month_locked",
        "month_unlocked",
    ]
    assert [row["input_revision"] for row in audit_rows] == [revision, revision]


def test_monthly_actual_costs_store_authoritative_totals_with_provenance(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    with db.connection() as connection:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        connection.execute(
            "INSERT INTO monthly_actual_costs"
            "(report_month, destination_id, cost_won, source_type, source_note) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", destination_id, 1200, "legacy", "confirmed ledger"),
        )
        row = connection.execute(
            "SELECT typeof(cost_won), source_type, source_note, created_at, updated_at "
            "FROM monthly_actual_costs"
        ).fetchone()
        assert tuple(row[:3]) == ("integer", "legacy", "confirmed ledger")
        assert row["created_at"]
        assert row["updated_at"]

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_actual_costs"
                "(report_month, destination_id, cost_won, source_type) "
                "VALUES (?, ?, ?, ?)",
                ("2026-08", destination_id, 1300, "finalized_import"),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_actual_costs"
                "(report_month, destination_id, cost_won, source_type) "
                "VALUES (?, ?, ?, ?)",
                ("2026-09", destination_id, "not-integer", "legacy"),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_actual_costs"
                "(report_month, destination_id, cost_won, source_type) "
                "VALUES (?, ?, ?, ?)",
                ("2026-09", destination_id, -1, "legacy"),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute("DELETE FROM destinations WHERE id = ?", (destination_id,))


def test_monthly_actual_costs_require_nonblank_source_type(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    with db.connection() as connection:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        sql = (
            "INSERT INTO monthly_actual_costs"
            "(report_month, destination_id, cost_won, source_type) "
            "VALUES (?, ?, ?, ?)"
        )
        connection.execute(
            sql,
            ("2026-08", destination_id, 1200, "legacy"),
        )

        for report_month, source_type in (("2026-09", ""), ("2026-10", "   ")):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    sql,
                    (report_month, destination_id, 1300, source_type),
                )


def test_transport_entry_month_must_match_its_import_batch(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, source_row, "
            "transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (batch_id, "2026-08", destination_id, "Sheet1", 2, 1, "regular", "1", 1000),
        )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_sheet, source_row, "
                "transport_day, transport_type, trip_count_text, cost_won) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (batch_id, "2026-09", destination_id, "Sheet1", 3, 1, "regular", "1", 1000),
            )
    finally:
        connection.close()


@pytest.mark.parametrize(
    ("existing_from", "existing_to", "new_from", "new_to", "overlaps"),
    [
        ("2026-03", "2026-06", "2026-07", "2026-08", False),
        ("2026-03", "2026-06", "2026-06", "2026-08", True),
        ("2026-03", "2026-06", "2026-04", "2026-05", True),
        ("2026-03", "2026-06", "2026-01", "2026-08", True),
        ("2026-03", None, "2027-01", None, True),
    ],
)
def test_vehicle_rate_inserts_reject_overlapping_effective_periods(
    tmp_path, existing_from, existing_to, new_from, new_to, overlaps
):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        connection.execute(
            "INSERT INTO vehicle_rates"
            "(destination_id, vehicle_type, unit_rate_won, effective_from, effective_to) "
            "VALUES (?, ?, ?, ?, ?)",
            (destination_id, "5-ton", 1000, existing_from, existing_to),
        )
        parameters = (destination_id, "5-ton", 1100, new_from, new_to)
        sql = (
            "INSERT INTO vehicle_rates"
            "(destination_id, vehicle_type, unit_rate_won, effective_from, effective_to) "
            "VALUES (?, ?, ?, ?, ?)"
        )

        if overlaps:
            with pytest.raises(sqlite3.IntegrityError, match="overlap"):
                connection.execute(sql, parameters)
        else:
            connection.execute(sql, parameters)
    finally:
        connection.close()


@pytest.mark.parametrize(
    ("existing_to", "updated_from", "updated_to", "overlaps"),
    [
        ("2026-06", "2026-07", "2026-08", False),
        ("2026-06", "2026-06", "2026-08", True),
        ("2026-06", "2026-04", "2026-05", True),
        ("2026-06", "2026-01", "2026-08", True),
        (None, "2027-03", None, True),
    ],
)
def test_vehicle_rate_updates_reject_overlapping_effective_periods(
    tmp_path, existing_to, updated_from, updated_to, overlaps
):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        connection.execute(
            "INSERT INTO vehicle_rates"
            "(destination_id, vehicle_type, unit_rate_won, effective_from, effective_to) "
            "VALUES (?, ?, ?, ?, ?)",
            (destination_id, "5-ton", 1000, "2026-03", existing_to),
        )
        rate_id = connection.execute(
            "INSERT INTO vehicle_rates"
            "(destination_id, vehicle_type, unit_rate_won, effective_from, effective_to) "
            "VALUES (?, ?, ?, ?, ?)",
            (destination_id, "5-ton", 1100, "2025-01", "2025-02"),
        ).lastrowid
        sql = (
            "UPDATE vehicle_rates SET effective_from = ?, effective_to = ? WHERE id = ?"
        )

        if overlaps:
            with pytest.raises(sqlite3.IntegrityError, match="overlap"):
                connection.execute(sql, (updated_from, updated_to, rate_id))
        else:
            connection.execute(sql, (updated_from, updated_to, rate_id))
    finally:
        connection.close()


def test_structural_integer_columns_reject_text_and_out_of_range_values(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        for explicit_id in (-1, "not-an-id"):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    "INSERT INTO destinations(id, name, display_order) VALUES (?, ?, ?)",
                    (explicit_id, f"Destination {explicit_id}", 1),
                )
        for display_order in (-1, "first"):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
                    (f"Destination {display_order}", display_order),
                )

        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid
        entry_sql = (
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, source_row, "
            "transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        )
        for source_row, transport_day in (("row", 1), (-1, 1), (2, "day"), (2, 0)):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    entry_sql,
                    (
                        batch_id,
                        "2026-08",
                        destination_id,
                        "Sheet1",
                        source_row,
                        transport_day,
                        "regular",
                        "1",
                        1000,
                    ),
                )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO report_runs(report_month, status, is_locked) VALUES (?, ?, ?)",
                ("2026-08", "draft", "locked"),
            )
    finally:
        connection.close()


def test_import_batch_sha256_requires_exactly_64_hexadecimal_characters(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        sql = (
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)"
        )
        connection.execute(
            sql,
            ("2026-08", "test", "lower.xlsx", "abcdef0123456789" * 4, "staged"),
        )
        connection.execute(
            sql,
            ("2026-09", "test", "upper.xlsx", "ABCDEF0123456789" * 4, "staged"),
        )
        for report_month, invalid_hash in (
            ("2026-10", "a" * 63),
            ("2026-11", "a" * 65),
            ("2026-12", "g" * 64),
        ):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    sql,
                    (report_month, "test", "invalid.xlsx", invalid_hash, "staged"),
                )
    finally:
        connection.close()


def test_migrations_are_repeatable(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    before = _schema_objects(db)

    db.migrate()

    assert _schema_objects(db) == before
    connection = db.connect()
    try:
        versions = connection.execute(
            "SELECT version FROM schema_migrations ORDER BY version"
        ).fetchall()
    finally:
        connection.close()
    assert [row["version"] for row in versions] == [1, 2, 3, 4, 5]


def test_migrate_rejects_schema_versions_newer_than_bundled_migrations(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    with db.connection() as connection:
        connection.execute(
            "INSERT INTO schema_migrations(version, applied_at) "
            "VALUES (?, datetime('now'))",
            (999,),
        )
        connection.commit()

    with pytest.raises(RuntimeError, match="newer.*bundled migrations"):
        db.migrate()


def test_report_group_members_store_a_valid_display_order(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        group_id = connection.execute(
            "INSERT INTO report_groups(name, display_order) VALUES (?, ?)",
            ("Region A", 1),
        ).lastrowid
        destination_ids = [
            connection.execute(
                "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
                (name, order),
            ).lastrowid
            for order, name in enumerate(("Destination A", "Destination B"), start=1)
        ]
        connection.execute(
            "INSERT INTO report_group_members"
            "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
            (group_id, destination_ids[0], 2),
        )
        connection.execute(
            "INSERT INTO report_group_members"
            "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
            (group_id, destination_ids[1], 1),
        )

        ordered_ids = connection.execute(
            "SELECT destination_id FROM report_group_members "
            "WHERE group_id = ? ORDER BY display_order",
            (group_id,),
        ).fetchall()
    finally:
        connection.close()

    assert [row["destination_id"] for row in ordered_ids] == [
        destination_ids[1],
        destination_ids[0],
    ]


def test_report_group_member_order_is_unique_within_each_group(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        group_ids = [
            connection.execute(
                "INSERT INTO report_groups(name, display_order) VALUES (?, ?)",
                (name, order),
            ).lastrowid
            for order, name in enumerate(("Region A", "Region B"), start=1)
        ]
        destination_ids = [
            connection.execute(
                "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
                (name, order),
            ).lastrowid
            for order, name in enumerate(("Destination A", "Destination B"), start=1)
        ]
        connection.execute(
            "INSERT INTO report_group_members"
            "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
            (group_ids[0], destination_ids[0], 1),
        )

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO report_group_members"
                "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
                (group_ids[0], destination_ids[1], 1),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO report_group_members"
                "(group_id, destination_id) VALUES (?, ?)",
                (group_ids[0], destination_ids[1]),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO report_group_members"
                "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
                (group_ids[0], destination_ids[1], -1),
            )

        connection.execute(
            "INSERT INTO report_group_members"
            "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
            (group_ids[1], destination_ids[1], 1),
        )
    finally:
        connection.close()


def test_failed_migration_rolls_back_schema_and_version(tmp_path):
    migrations = tmp_path / "migrations"
    migrations.mkdir()
    default_migrations = Database(tmp_path / "unused.db").migrations_dir
    shutil.copyfile(
        default_migrations / "001_initial.sql",
        migrations / "001_initial.sql",
    )
    (migrations / "002_broken.sql").write_text(
        "CREATE TABLE transient_table(id INTEGER PRIMARY KEY);\n"
        "INSERT INTO table_that_does_not_exist(id) VALUES (1);\n",
        encoding="utf-8",
    )
    db = Database(tmp_path / "app.db", migrations_dir=migrations)

    with pytest.raises(sqlite3.OperationalError):
        db.migrate()

    assert "transient_table" not in db.table_names()
    connection = db.connect()
    try:
        versions = connection.execute(
            "SELECT version FROM schema_migrations ORDER BY version"
        ).fetchall()
    finally:
        connection.close()
    assert [row["version"] for row in versions] == [1]


def test_connections_enable_and_enforce_foreign_keys(tmp_path):
    db = Database(tmp_path / "nested" / "app.db")
    db.migrate()

    connection = db.connect()
    try:
        assert connection.execute("PRAGMA foreign_keys").fetchone()[0] == 1
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO destination_aliases"
                "(raw_name, source_type, destination_id) VALUES (?, ?, ?)",
                ("unknown", "test", 999),
            )
    finally:
        connection.close()


def test_connection_context_manager_closes_after_an_exception(tmp_path):
    db = Database(tmp_path / "app.db")
    owned_connection = None

    with pytest.raises(RuntimeError, match="caller failed"):
        with db.connection() as connection:
            owned_connection = connection
            connection.execute("SELECT 1")
            raise RuntimeError("caller failed")

    assert owned_connection is not None
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        owned_connection.execute("SELECT 1")


def test_schema_includes_required_business_and_audit_fields(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()

    expected_columns = {
        "destinations": {"representative_item"},
        "monthly_plans": {
            "quantity_ea_text",
            "cost_won",
            "representative_item",
            "source_note",
        },
        "monthly_sales": {"amount_won", "source_note", "confirmed_at"},
        "import_batches": {
            "file_sha256",
            "status",
            "error_summary",
            "is_current",
        },
        "transport_entries": {
            "source_sheet",
            "source_row",
            "source_alias",
            "source_date",
            "transport_day",
            "transport_type",
            "trip_count_text",
            "unit_rate_won",
            "cost_won",
            "unresolved_alias",
        },
        "report_runs": {
            "is_locked",
            "locked_at",
            "unlocked_at",
            "unlock_reason",
            "input_revision",
            "output_hashes_json",
        },
    }
    connection = db.connect()
    try:
        actual_columns = {
            table: {
                row["name"]
                for row in connection.execute(f"PRAGMA table_info({table})")
            }
            for table in expected_columns
        }
    finally:
        connection.close()

    for table, required in expected_columns.items():
        assert required <= actual_columns[table]


@pytest.mark.parametrize(
    "invalid_source_date",
    ["2026-8-01", "2026/08/01", "2026-02-30", "not-a-date", ""],
)
def test_transport_source_date_requires_valid_iso_date_when_present(
    tmp_path, invalid_source_date
):
    db = Database(tmp_path / "app.db")
    db.migrate()
    with db.connection() as connection:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "transport", "source.xlsx", "d" * 64, "imported"),
        ).lastrowid
        sql = (
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, source_date, transport_day, transport_type, "
            "trip_count_text, cost_won) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                sql,
                (
                    batch_id,
                    "2026-08",
                    destination_id,
                    "Sheet1",
                    2,
                    invalid_source_date,
                    1,
                    "regular",
                    "1",
                    1000,
                ),
            )
        connection.execute(
            sql,
            (
                batch_id,
                "2026-08",
                destination_id,
                "Sheet1",
                3,
                "2026-08-01",
                1,
                "regular",
                "1",
                1000,
            ),
        )
        connection.execute(
            sql,
            (
                batch_id,
                "2026-08",
                destination_id,
                "Sheet1",
                4,
                None,
                1,
                "regular",
                "1",
                1000,
            ),
        )


def test_monetary_values_require_integer_won(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_sales(report_month, amount_won) VALUES (?, ?)",
                ("2026-08", 1.5),
            )
    finally:
        connection.close()


def test_monthly_records_enforce_month_uniqueness_and_business_references(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        connection.execute(
            "INSERT INTO monthly_plans"
            "(report_month, destination_id, quantity_ea_text, cost_won) "
            "VALUES (?, ?, ?, ?)",
            ("2026-08", destination_id, "1.25", 1000),
        )
        stored_types = connection.execute(
            "SELECT typeof(quantity_ea_text), typeof(cost_won) FROM monthly_plans"
        ).fetchone()
        assert tuple(stored_types) == ("text", "integer")

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_plans"
                "(report_month, destination_id, quantity_ea_text, cost_won) "
                "VALUES (?, ?, ?, ?)",
                ("2026-08", destination_id, "2", 2000),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_actual_quantities"
                "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
                ("2026-13", destination_id, "1"),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "DELETE FROM destinations WHERE id = ?",
                (destination_id,),
            )
    finally:
        connection.close()


def test_migrations_run_in_numeric_version_order(tmp_path):
    migrations = tmp_path / "migrations"
    migrations.mkdir()
    (migrations / "001_initial.sql").write_text(
        "CREATE TABLE schema_migrations("
        "version INTEGER PRIMARY KEY, applied_at TEXT NOT NULL);\n"
        "CREATE TABLE migration_order(position INTEGER PRIMARY KEY, version INTEGER);",
        encoding="utf-8",
    )
    (migrations / "010_tenth.sql").write_text(
        "INSERT INTO migration_order(version) VALUES (10);",
        encoding="utf-8",
    )
    (migrations / "002_second.sql").write_text(
        "INSERT INTO migration_order(version) VALUES (2);",
        encoding="utf-8",
    )
    db = Database(tmp_path / "app.db", migrations_dir=migrations)

    db.migrate()

    connection = db.connect()
    try:
        versions = connection.execute(
            "SELECT version FROM migration_order ORDER BY position"
        ).fetchall()
    finally:
        connection.close()
    assert [row["version"] for row in versions] == [2, 10]


def test_unresolved_entries_and_unlocks_require_explanations(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid
        for source_row, unresolved_alias in enumerate((None, "  "), start=2):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    "INSERT INTO transport_entries"
                    "(import_batch_id, report_month, unresolved_alias, source_sheet, "
                    "source_row, transport_day, transport_type, trip_count_text, "
                    "cost_won) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        batch_id,
                        "2026-08",
                        unresolved_alias,
                        "Sheet1",
                        source_row,
                        1,
                        "regular",
                        "1",
                        1000,
                    ),
                )
        for unlock_reason in (None, "  "):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(
                    "INSERT INTO report_runs"
                    "(report_month, status, unlocked_at, unlock_reason) "
                    "VALUES (?, ?, ?, ?)",
                    (
                        "2026-08",
                        "unlocked",
                        "2026-09-15T10:00:00",
                        unlock_reason,
                    ),
                )
    finally:
        connection.close()


def test_quantity_and_trip_values_require_nonnegative_decimal_text(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid

        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_plans"
                "(report_month, destination_id, quantity_ea_text, cost_won) "
                "VALUES (?, ?, ?, ?)",
                ("2026-08", destination_id, "not-a-number", 1000),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO monthly_actual_quantities"
                "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
                ("2026-08", destination_id, "1,25"),
            )
        with pytest.raises(sqlite3.IntegrityError):
            connection.execute(
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_sheet, "
                "source_row, transport_day, transport_type, trip_count_text, cost_won) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    batch_id,
                    "2026-08",
                    destination_id,
                    "Sheet1",
                    2,
                    1,
                    "regular",
                    "-1",
                    1000,
                ),
            )

        connection.execute(
            "INSERT INTO monthly_actual_quantities"
            "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
            ("2026-08", destination_id, "0"),
        )
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-08",
                destination_id,
                "Sheet1",
                3,
                2,
                "regular",
                "1.25",
                1250,
            ),
        )
    finally:
        connection.close()


@pytest.mark.parametrize(
    "invalid_value",
    ["01", "01.20", "1.", ".5", "00", "0.0", "1.0", "1.20", "1.250"],
)
def test_decimal_text_columns_reject_noncanonical_values(tmp_path, invalid_value):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid
        statements = [
            (
                "INSERT INTO monthly_plans"
                "(report_month, destination_id, quantity_ea_text, cost_won) "
                "VALUES (?, ?, ?, ?)",
                ("2026-08", destination_id, invalid_value, 1000),
            ),
            (
                "INSERT INTO monthly_actual_quantities"
                "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
                ("2026-08", destination_id, invalid_value),
            ),
            (
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_sheet, "
                "source_row, transport_day, transport_type, trip_count_text, cost_won) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    batch_id,
                    "2026-08",
                    destination_id,
                    "Sheet1",
                    2,
                    1,
                    "regular",
                    invalid_value,
                    1000,
                ),
            ),
        ]
        for sql, parameters in statements:
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(sql, parameters)
    finally:
        connection.close()


@pytest.mark.parametrize("valid_value", ["0", "1", "125", "0.5", "1.25"])
def test_decimal_text_columns_accept_canonical_values(tmp_path, valid_value):
    db = Database(tmp_path / "app.db")
    db.migrate()
    connection = db.connect()
    try:
        destination_id = connection.execute(
            "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
            ("Destination A", 1),
        ).lastrowid
        batch_id = connection.execute(
            "INSERT INTO import_batches"
            "(report_month, source_type, source_filename, file_sha256, status) "
            "VALUES (?, ?, ?, ?, ?)",
            ("2026-08", "test", "source.xlsx", "a" * 64, "staged"),
        ).lastrowid

        connection.execute(
            "INSERT INTO monthly_plans"
            "(report_month, destination_id, quantity_ea_text, cost_won) "
            "VALUES (?, ?, ?, ?)",
            ("2026-08", destination_id, valid_value, 1000),
        )
        connection.execute(
            "INSERT INTO monthly_actual_quantities"
            "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
            ("2026-08", destination_id, valid_value),
        )
        connection.execute(
            "INSERT INTO transport_entries"
            "(import_batch_id, report_month, destination_id, source_sheet, "
            "source_row, transport_day, transport_type, trip_count_text, cost_won) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                batch_id,
                "2026-08",
                destination_id,
                "Sheet1",
                2,
                1,
                "regular",
                valid_value,
                1000,
            ),
        )
    finally:
        connection.close()


def test_migration_resources_must_exist_and_be_numbered(tmp_path):
    missing = tmp_path / "missing"
    with pytest.raises(FileNotFoundError):
        Database(tmp_path / "missing.db", migrations_dir=missing).migrate()

    empty = tmp_path / "empty"
    empty.mkdir()
    with pytest.raises(RuntimeError, match="No migrations found"):
        Database(tmp_path / "empty.db", migrations_dir=empty).migrate()

    invalid = tmp_path / "invalid"
    invalid.mkdir()
    (invalid / "initial.sql").write_text("SELECT 1;", encoding="utf-8")
    with pytest.raises(ValueError, match="Invalid migration filename"):
        Database(tmp_path / "invalid.db", migrations_dir=invalid).migrate()


def test_concurrent_migrate_calls_do_not_reapply_versions(tmp_path, monkeypatch):
    barrier = threading.Barrier(2)
    original_connect = sqlite3.connect

    class SynchronizedConnection(sqlite3.Connection):
        def execute(self, sql, parameters=()):
            if sql == "BEGIN IMMEDIATE":
                barrier.wait(timeout=5)
            return super().execute(sql, parameters)

    def synchronized_connect(*args, **kwargs):
        return original_connect(*args, factory=SynchronizedConnection, **kwargs)

    monkeypatch.setattr(db_module.sqlite3, "connect", synchronized_connect)
    database_path = tmp_path / "app.db"

    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = [
            executor.submit(Database(database_path).migrate)
            for _ in range(2)
        ]
        for future in futures:
            future.result()

    connection = Database(database_path).connect()
    try:
        versions = connection.execute(
            "SELECT version FROM schema_migrations ORDER BY version"
        ).fetchall()
    finally:
        connection.close()
    assert [row["version"] for row in versions] == [1, 2, 3, 4, 5]


def _schema_objects(db):
    connection = db.connect()
    try:
        return connection.execute(
            "SELECT type, name, sql FROM sqlite_master "
            "WHERE name NOT LIKE 'sqlite_%' ORDER BY type, name"
        ).fetchall()
    finally:
        connection.close()
