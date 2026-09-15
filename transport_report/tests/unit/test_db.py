import shutil
import sqlite3
import threading
from concurrent.futures import ThreadPoolExecutor

import pytest

import app.db as db_module
from app.db import Database


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
        "monthly_sales",
        "transport_entries",
        "import_batches",
        "report_runs",
    } <= names


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
    assert [row["version"] for row in versions] == [1]


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
        },
        "transport_entries": {
            "source_sheet",
            "source_row",
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
    assert [row["version"] for row in versions] == [1]


def _schema_objects(db):
    connection = db.connect()
    try:
        return connection.execute(
            "SELECT type, name, sql FROM sqlite_master "
            "WHERE name NOT LIKE 'sqlite_%' ORDER BY type, name"
        ).fetchall()
    finally:
        connection.close()
