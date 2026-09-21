from decimal import Decimal
import shutil
import sqlite3

import pytest

from app.db import Database
from app.repositories.masters import MasterRepository
from app.repositories.monthly_inputs import (
    MonthlyInputRepository, MonthlyInputValidationError, MonthLockedError,
    MonthlyInputRevisionError,
)


@pytest.fixture
def db(tmp_path):
    db = Database(tmp_path / "inputs.db")
    db.migrate()
    return db


def test_supplemental_migration_upgrade_preserves_missing_sales(tmp_path):
    bundled = Database(tmp_path / "unused.db").migrations_dir
    old = tmp_path / "old"
    old.mkdir()
    for source in bundled.glob("00[1-5]_*.sql"):
        shutil.copyfile(source, old / source.name)
    db = Database(tmp_path / "upgrade.db", old)
    db.migrate()
    MonthlyInputRepository(db).save_sales("2026-08", 0, "confirmed", "2026-08-31")
    upgraded = Database(db.path)
    upgraded.migrate()
    upgraded.migrate()
    assert "report_supplemental_inputs" in upgraded.table_names()
    assert MonthlyInputRepository(upgraded).get_sales("2026-08").amount_won == 0
    with upgraded.connection() as connection:
        assert connection.execute("SELECT count(*) FROM report_supplemental_inputs").fetchone()[0] == 0


def test_missing_zero_roundtrip_and_pair_validation(db):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    repo = ReportInputRepository(db)
    assert repo.get("2026-08") is None
    record = repo.save(ReportSupplementalInputs("2026-08", planned_sales_won=0,
        nonregular_actual_quantity_ea=Decimal("0.00"), nonregular_actual_cost_won=0,
        source_note=" 수기 확인 "))
    assert record.planned_sales_won == 0
    assert record.nonregular_actual_quantity_ea == 0
    assert record.nonregular_planned_quantity_ea is None
    assert record.source_note == "수기 확인"
    assert record.updated_at
    assert MonthlyInputRepository(db).get_sales("2026-08") is None
    with pytest.raises(MonthlyInputValidationError):
        repo.save(ReportSupplementalInputs("2026-08", nonregular_actual_cost_won=1))
    assert repo.get("2026-08") == record
    cleared = repo.save(ReportSupplementalInputs("2026-08"))
    assert cleared.planned_sales_won is None


@pytest.mark.parametrize("field,value", [
    ("planned_sales_won", -1), ("planned_sales_won", True),
    ("planned_sales_won", 2**63), ("nonregular_actual_quantity_ea", Decimal("NaN")),
])
def test_invalid_supplemental_values_rejected(db, field, value):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    with pytest.raises(MonthlyInputValidationError):
        ReportInputRepository(db).save(ReportSupplementalInputs("2026-08", **{field: value}))


def test_supplemental_database_constraints_and_locks(db):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    repo = ReportInputRepository(db)
    repo.save(ReportSupplementalInputs("2026-08"))
    with db.connection() as connection:
        for sql in (
            "UPDATE report_supplemental_inputs SET nonregular_actual_cost_won = 0",
            "UPDATE report_supplemental_inputs SET planned_sales_won = -1",
            "UPDATE report_supplemental_inputs SET report_month = '2026-13'",
            "UPDATE report_supplemental_inputs SET nonregular_actual_quantity_text = '01', nonregular_actual_cost_won = 0",
        ):
            with pytest.raises(sqlite3.IntegrityError):
                connection.execute(sql)
    MonthlyInputRepository(db).finalize_month("2026-08", "revision")
    with pytest.raises(MonthLockedError):
        repo.save(ReportSupplementalInputs("2026-08", planned_sales_won=10))
    with db.connection() as connection:
        for sql in (
            "DELETE FROM report_supplemental_inputs",
            "UPDATE report_supplemental_inputs SET source_note = 'changed'",
            "INSERT OR REPLACE INTO report_supplemental_inputs(report_month) VALUES ('2026-08')",
        ):
            with pytest.raises(sqlite3.IntegrityError, match="locked"):
                connection.execute(sql)


def test_form_is_atomic_and_rejects_stale_next_plan(db):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    dest = MasterRepository(db).create_destination("납품처", 1)
    repo = ReportInputRepository(db)
    snapshot = repo.get_form_snapshot("2026-08")
    rows = [dict(destination_id=dest.id, plan_quantity=Decimal(0), plan_cost=0, representative_item="제품")]
    repo.save_form("2026-08", ReportSupplementalInputs("2026-08", nonregular_actual_quantity_ea=Decimal(0), nonregular_actual_cost_won=0),
        ReportSupplementalInputs("2026-09", planned_sales_won=0), rows, expected_revision=snapshot.revision)
    assert MonthlyInputRepository(db).get_plan("2026-09", dest.id).cost_won == 0
    with pytest.raises(MonthlyInputRevisionError):
        repo.save_form("2026-08", ReportSupplementalInputs("2026-08", planned_sales_won=999), ReportSupplementalInputs("2026-09"), rows, expected_revision=snapshot.revision)
    assert repo.get("2026-08").planned_sales_won is None
    snapshot = repo.get_form_snapshot("2026-08")
    MonthlyInputRepository(db).finalize_month("2026-09", "lock")
    with pytest.raises(MonthLockedError):
        repo.save_form("2026-08", ReportSupplementalInputs("2026-08", planned_sales_won=999), ReportSupplementalInputs("2026-09"), rows, expected_revision=snapshot.revision)
    assert repo.get("2026-08").planned_sales_won is None


def test_supplemental_migration_failure_rolls_back(tmp_path):
    bundled = Database(tmp_path / "unused.db").migrations_dir
    migrations = tmp_path / "migrations"
    shutil.copytree(bundled, migrations)
    migration = migrations / "006_report_supplemental_inputs.sql"
    assert migration.exists()
    migration.write_text(migration.read_text(encoding="utf-8") + "\nSELECT missing_column FROM missing_table;", encoding="utf-8")
    db = Database(tmp_path / "rollback.db", migrations)
    with pytest.raises(sqlite3.OperationalError):
        db.migrate()
    assert "report_supplemental_inputs" not in db.table_names()
    with db.connection() as connection:
        assert connection.execute("SELECT max(version) FROM schema_migrations").fetchone()[0] == 5


def test_late_plan_failure_rolls_back_both_supplemental_months(db):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    dest = MasterRepository(db).create_destination("납품처", 1)
    repo = ReportInputRepository(db)
    snapshot = repo.get_form_snapshot("2026-08")
    with db.connection() as connection, connection:
        connection.execute("CREATE TRIGGER fail_next_plan BEFORE INSERT ON monthly_plans BEGIN SELECT RAISE(ABORT, 'forced write failure'); END;")
    with pytest.raises(sqlite3.IntegrityError, match="forced write failure"):
        repo.save_form("2026-08", ReportSupplementalInputs("2026-08", planned_sales_won=1),
            ReportSupplementalInputs("2026-09", planned_sales_won=2),
            [dict(destination_id=dest.id, plan_quantity=0, plan_cost=0, representative_item=None)], expected_revision=snapshot.revision)
    assert repo.get("2026-08") is None
    assert repo.get("2026-09") is None
    assert MonthlyInputRepository(db).get_plan("2026-09", dest.id) is None


def test_form_current_lock_and_master_change_rejected(db):
    from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs
    repo = ReportInputRepository(db)
    snapshot = repo.get_form_snapshot("2026-08")
    MasterRepository(db).create_destination("new", 1)
    with pytest.raises(MonthlyInputRevisionError):
        repo.save_form("2026-08", ReportSupplementalInputs("2026-08"), ReportSupplementalInputs("2026-09"), [], expected_revision=snapshot.revision)
    MonthlyInputRepository(db).finalize_month("2026-08", "locked")
    with pytest.raises(MonthLockedError):
        repo.save_form("2026-08", ReportSupplementalInputs("2026-08"), ReportSupplementalInputs("2026-09"), [], expected_revision=snapshot.revision)
