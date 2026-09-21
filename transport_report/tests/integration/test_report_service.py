"""Task13 uses fictional persisted inputs and a synthetic template only."""
from datetime import datetime
from hashlib import sha256
import importlib
import json
from pathlib import Path
import sqlite3

import pytest

from app.config import RuntimePaths
from app.db import Database
from tests.integration.test_pptx_report import build_template


NOW = datetime(2026, 9, 2, 10, 11, 12)


def service_module():
    try:
        return importlib.import_module("app.services.report_service")
    except ModuleNotFoundError:
        pytest.fail("Task13 report orchestration service is not implemented")


def backup_module():
    try:
        return importlib.import_module("app.services.backup_service")
    except ModuleNotFoundError:
        pytest.fail("Task13 SQLite backup and restore service is not implemented")


def test_service_contract_exists():
    assert callable(service_module().ReportService)
    assert callable(backup_module().BackupService)


@pytest.fixture
def prepared(tmp_path):
    paths = RuntimePaths.from_root(tmp_path)
    paths.ensure()
    paths.template.parent.mkdir()
    build_template(paths.template)
    db = Database(paths.database)
    db.migrate()
    months = [*(f"2025-{m:02}" for m in range(1, 13)), *(f"2026-{m:02}" for m in range(1, 10))]
    with db.connection() as c:
        for i in range(1, 13):
            c.execute("INSERT INTO destinations(id,name,display_order) VALUES (?,?,?)", (i, f"가상{i:02}", i))
            c.execute("INSERT INTO destination_aliases(raw_name,source_type,destination_id) VALUES (?,?,?)", (f"원천{i:02}", "transport", i))
            for month in months:
                c.execute("INSERT INTO monthly_plans(report_month,destination_id,quantity_ea_text,cost_won) VALUES (?,?,?,?)", (month, i, "5000", 2000000))
                if month <= "2026-08":
                    c.execute("INSERT INTO monthly_actual_quantities(report_month,destination_id,quantity_ea_text) VALUES (?,?,?)", (month, i, "5000"))
                    c.execute("INSERT INTO monthly_actual_costs(report_month,destination_id,cost_won,source_type) VALUES (?,?,?,?)", (month, i, 2100000, "manual"))
        for month in months:
            c.execute("INSERT INTO monthly_sales(report_month,amount_won,confirmed_at) VALUES (?,?,?)", (month, 310000000, month + "-28"))
            c.execute("INSERT INTO report_supplemental_inputs(report_month,planned_sales_won,nonregular_planned_quantity_text,nonregular_planned_cost_won,nonregular_actual_quantity_text,nonregular_actual_cost_won,source_note) VALUES (?,?,?,?,?,?,?)", (month, 300000000, "0", 0, "0", 0, "가상 확인 입력"))
        c.execute("INSERT INTO import_batches(id,report_month,source_type,source_filename,file_sha256,status,is_current) VALUES (1,'2026-08','transport','fiction.xlsx',?,'imported',1)", ("a" * 64,))
        for i in range(1, 13):
            c.execute("INSERT INTO transport_entries(import_batch_id,report_month,destination_id,source_sheet,source_row,transport_day,transport_type,trip_count_text,cost_won,source_alias) VALUES (1,'2026-08',?,'가상',?,1,'regular','1',2100000,?)", (i, i, f"원천{i:02}"))
        c.commit()
    return db, paths


@pytest.mark.parametrize("sql,code", [
    ("DELETE FROM monthly_sales WHERE report_month='2026-08'", "MISSING_SALES"),
    ("DELETE FROM monthly_actual_quantities WHERE report_month='2026-08' AND destination_id=1", "MISSING_ACTUAL_QUANTITY"),
    ("DELETE FROM destination_aliases WHERE destination_id=1", "MISSING_DESTINATION_ALIAS"),
    ("DELETE FROM monthly_actual_quantities WHERE report_month='2025-01' AND destination_id=1", "MISSING_PRIOR_YEAR_HISTORY"),
    ("DELETE FROM monthly_plans WHERE report_month='2026-09' AND destination_id=1", "MISSING_NEXT_MONTH_PLAN"),
])
def test_prerequisites_block_without_partial_deliverables(prepared, sql, code):
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute(sql)
        c.commit()
    with pytest.raises(module.ReportGenerationError) as failure:
        module.ReportService(db, paths, clock=lambda: NOW).generate("2026-08")
    assert code in {issue.code for issue in failure.value.issues}
    assert not list(paths.outputs.rglob("*.*"))
    with db.connection() as c:
        assert c.execute("SELECT COUNT(*) FROM report_runs").fetchone()[0] == 0


def test_complete_output_hashes_snapshot_and_template(prepared):
    module = service_module()
    db, paths = prepared
    before = sha256(paths.template.read_bytes()).hexdigest()
    result = module.ReportService(db, paths, clock=lambda: NOW).generate("2026-08")
    assert result.directory == paths.outputs / "2026-08" / "run-20260902-101112"
    assert sorted(p.suffix for p in result.directory.iterdir()) == [".png", ".png", ".png", ".pptx", ".xlsx"]
    with db.connection() as c:
        audit = c.execute("SELECT * FROM report_runs").fetchone()
    manifest = json.loads(audit["output_hashes_json"])
    assert audit["input_revision"] == sha256(manifest["input_snapshot_json"].encode()).hexdigest()
    assert manifest["model_sha256"] == sha256(manifest["model_json"].encode()).hexdigest()
    assert json.loads(manifest["model_json"])["report"]["next_month"]["month"] == "2026-09"
    for item in manifest["files"]:
        assert not Path(item["path"]).is_absolute()
        file = paths.outputs / item["path"]
        assert item["sha256"] == sha256(file.read_bytes()).hexdigest()
        assert item["size"] == file.stat().st_size
    assert sha256(paths.template.read_bytes()).hexdigest() == before


@pytest.mark.parametrize("stage", ["charts", "xlsx", "pptx", "verify", "rename", "audit"])
def test_failure_at_each_stage_is_atomic(prepared, monkeypatch, stage):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    prior = paths.outputs / "2026-08" / "prior"
    prior.mkdir(parents=True)
    (prior / "keep.txt").write_text("existing", encoding="utf-8")
    def fail(*args):
        raise OSError("injected internal path D:/private/data")
    monkeypatch.setattr(service, "_" + stage, fail)
    with pytest.raises(module.ReportGenerationError) as failure:
        service.generate("2026-08")
    assert "D:/private" not in str(failure.value)
    assert list((paths.outputs / "2026-08").iterdir()) == [prior]
    assert (prior / "keep.txt").read_text() == "existing"
    with db.connection() as c:
        assert c.execute("SELECT COUNT(*) FROM report_runs").fetchone()[0] == 0


def test_same_timestamp_separate_runs_and_snapshot_survives_concurrent_edit(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    first = service.generate("2026-08")
    original = service._charts
    def edit_during_render(*args):
        with db.connection() as c:
            c.execute("UPDATE monthly_sales SET amount_won=999999999 WHERE report_month='2026-08'")
            c.commit()
        original(*args)
    monkeypatch.setattr(service, "_charts", edit_during_render)
    second = service.generate("2026-08")
    assert first.directory != second.directory
    assert second.directory.name == "run-20260902-101112-001"
    with db.connection() as c:
        rows = c.execute("SELECT * FROM report_runs ORDER BY id").fetchall()
    assert rows[0]["input_revision"] == rows[1]["input_revision"]
    for row in rows:
        model = json.loads(json.loads(row["output_hashes_json"])["model_json"])
        assert model["report"]["sales"]["actual_won"] == 310000000


def test_backup_restore_prebackup_and_reject_bad_files(prepared, tmp_path):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    backup = service.backup()
    assert "20260902-101112" in backup.name and "v006" in backup.name
    assert service.backup() != backup
    with db.connection() as c:
        c.execute("UPDATE monthly_sales SET amount_won=777 WHERE report_month='2026-08'")
        c.commit()
    with pytest.raises(module.BackupError, match="연결"):
        service.restore(backup)
    result = service.restore(backup, exclusive_access=True)
    with db.connection() as c:
        assert c.execute("SELECT amount_won FROM monthly_sales WHERE report_month='2026-08'").fetchone()[0] == 310000000
    with sqlite3.connect(result.pre_restore_backup) as c:
        assert c.execute("SELECT amount_won FROM monthly_sales WHERE report_month='2026-08'").fetchone()[0] == 777
    before = db.path.read_bytes()
    for content in (b"not sqlite", b"SQLite format 3\x00" + b"bad" * 200):
        bad = tmp_path / "bad.db"
        bad.write_bytes(content)
        with pytest.raises(module.BackupError):
            service.restore(bad, exclusive_access=True)
        assert db.path.read_bytes() == before


@pytest.mark.parametrize("version", [5, 7])
def test_restore_rejects_inexact_schema(prepared, version):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    with sqlite3.connect(candidate) as c:
        if version == 5:
            c.execute("DELETE FROM schema_migrations WHERE version=6")
        else:
            c.execute("INSERT INTO schema_migrations VALUES (7,'future')")
    before = db.path.read_bytes()
    with pytest.raises(module.BackupError, match="버전"):
        service.restore(candidate, exclusive_access=True)
    assert db.path.read_bytes() == before


def test_restore_post_replacement_failure_rolls_back(prepared, monkeypatch):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    with db.connection() as c:
        c.execute("UPDATE monthly_sales SET amount_won=777 WHERE report_month='2026-08'")
        c.commit()
    monkeypatch.setattr(service, "_verify_restored", lambda: (_ for _ in ()).throw(OSError("failure")))
    with pytest.raises(module.BackupError):
        service.restore(candidate, exclusive_access=True)
    with db.connection() as c:
        assert c.execute("SELECT amount_won FROM monthly_sales WHERE report_month='2026-08'").fetchone()[0] == 777
    assert list(paths.backups.glob("pre-restore*"))


def test_snapshot_read_failure_has_korean_actionable_error(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths)
    monkeypatch.setattr(service, "_snapshot", lambda: (_ for _ in ()).throw(sqlite3.OperationalError("D:/secret database locked")))
    with pytest.raises(module.ReportGenerationError, match="다시"):
        service.generate()
    assert not list(paths.outputs.iterdir())


def test_exact_requested_run_collision_does_not_duplicate_audit(prepared):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    first = service.generate(requested_run_id="run-20260902-101112")
    files = {p.name: p.read_bytes() for p in first.directory.iterdir()}
    with pytest.raises(module.ReportGenerationError):
        service.generate(requested_run_id="run-20260902-101112")
    assert files == {p.name: p.read_bytes() for p in first.directory.iterdir()}
    with db.connection() as c:
        assert c.execute("SELECT COUNT(*) FROM report_runs").fetchone()[0] == 1


def test_audit_database_rejection_rolls_back_published_directory(prepared):
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("CREATE TRIGGER reject_report_audit BEFORE INSERT ON report_runs BEGIN SELECT RAISE(ABORT, 'test audit failure'); END")
        c.commit()
    with pytest.raises(module.ReportGenerationError):
        module.ReportService(db, paths, clock=lambda: NOW).generate()
    assert not list(paths.outputs.rglob("*.*"))
    with db.connection() as c:
        assert c.execute("SELECT COUNT(*) FROM report_runs").fetchone()[0] == 0


def test_restore_rejects_missing_application_table(prepared):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    with sqlite3.connect(candidate) as c:
        c.execute("DROP TABLE vehicle_rates")
    before = db.path.read_bytes()
    with pytest.raises(module.BackupError):
        service.restore(candidate, exclusive_access=True)
    assert db.path.read_bytes() == before
    assert not list(paths.backups.glob("pre-restore*"))


def test_restore_refuses_active_writer(prepared):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    with db.connection() as c:
        c.execute("BEGIN IMMEDIATE")
        with pytest.raises(module.BackupError, match="연결"):
            service.restore(candidate, exclusive_access=True)
        c.rollback()
    assert not list(paths.backups.glob("pre-restore*"))


def test_online_backup_includes_committed_wal_values(prepared):
    module = backup_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("UPDATE monthly_sales SET amount_won=456 WHERE report_month='2026-08'")
        c.commit()
        path = module.BackupService(db, paths.backups, clock=lambda: NOW).backup()
    with sqlite3.connect(path) as c:
        assert c.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        assert c.execute("SELECT amount_won FROM monthly_sales WHERE report_month='2026-08'").fetchone()[0] == 456


def test_mixed_manual_quantity_and_imported_cost_is_derived_evidence(prepared):
    module = service_module()
    db, paths = prepared
    bundle, _ = module.ReportService(db, paths).prepare()
    actuals = [o for o in bundle.operations if o.record_kind == "destination" and o.value_role == "actual"]
    assert len(actuals) == 12
    assert all(o.input_kind == "derived" for o in actuals)
    assert all("ERP" in o.source_locator and "1:" in o.source_locator for o in actuals)
    from app.reporting.xlsx_review import export_review_workbook
    from openpyxl import load_workbook
    export_review_workbook(bundle, paths.outputs / "evidence.xlsx")
    book = load_workbook(paths.outputs / "evidence.xlsx")
    try:
        assert sum(cell.value == "혼합 원천 집계" for row in book["운행실적"] for cell in row) == 12
    finally:
        book.close()


def test_generation_endpoint_produces_real_outputs_and_rechecks_gate(prepared):
    from app.main import create_app
    from tests.integration.test_web_routes import WebClient, post
    db, paths = prepared
    application = create_app(database=db, paths=paths)
    client = WebClient(application)
    response = client.get("/")
    assert 'data-can-generate="true"' in response.text
    response = post(client, "/months/2026-08/generate/excel", {"intent": "generate"})
    assert response.status_code == 200
    assert response.content[:2] == b"PK"
    assert len(list(paths.outputs.rglob("*.xlsx"))) == 1
    assert len(list(paths.outputs.rglob("*.pptx"))) == 1
    assert len(list(paths.outputs.rglob("*.png"))) == 3
    with db.connection() as c:
        c.execute("DELETE FROM monthly_sales WHERE report_month='2026-08'")
        c.commit()
    assert post(client, "/months/2026-08/generate/ppt", {"intent": "generate"}).status_code == 409
    assert len(list(paths.outputs.rglob("*.pptx"))) == 1


def test_failed_backup_closes_all_connections_and_removes_temporary(prepared, monkeypatch):
    module = backup_module()
    db, paths = prepared
    opened, closed = [], []
    original_connect = sqlite3.connect
    class FailingConnection(sqlite3.Connection):
        def backup(self, *args, **kwargs):
            raise sqlite3.OperationalError("injected backup failure")
        def close(self):
            closed.append(self)
            super().close()
    def connect(*args, **kwargs):
        connection = original_connect(*args, **kwargs, factory=FailingConnection)
        opened.append(connection)
        return connection
    monkeypatch.setattr(module.sqlite3, "connect", connect)
    with pytest.raises(module.BackupError):
        module.BackupService(db, paths.backups).backup()
    assert set(closed) == set(opened)
    assert not list(paths.backups.iterdir())


@pytest.mark.parametrize("statements", [
    ["DROP TRIGGER report_supplemental_inputs_reject_locked_update"],
    ["DROP INDEX uq_import_batches_current_month_source"],
    ["DROP INDEX uq_import_batches_current_month_source",
     "CREATE INDEX uq_import_batches_current_month_source ON import_batches(report_month, source_type) WHERE is_current = 1"],
    ["DROP TRIGGER report_supplemental_inputs_reject_locked_update",
     "CREATE TRIGGER report_supplemental_inputs_reject_locked_update BEFORE UPDATE ON report_supplemental_inputs WHEN 0 BEGIN SELECT RAISE(ABORT, 'report month locked'); END"],
    ["PRAGMA writable_schema=ON",
     "UPDATE sqlite_master SET sql=replace(sql, 'planned_sales_won >= 0', 'planned_sales_won >= -1') WHERE type='table' AND name='report_supplemental_inputs'",
     "PRAGMA writable_schema=OFF"],
    ["CREATE INDEX extra_sales_index ON monthly_sales(amount_won)"],
    ["CREATE VIEW extra_sales_view AS SELECT * FROM monthly_sales"],
], ids=["missing-lock-trigger", "missing-unique-index", "weakened-index", "weakened-lock-trigger", "weakened-table-check", "unexpected-index", "unexpected-view"])
def test_restore_rejects_changed_application_schema_before_touching_live(prepared, statements):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    connection = sqlite3.connect(candidate)
    try:
        for statement in statements:
            connection.execute(statement)
        connection.commit()
    finally:
        connection.close()
    before = db.path.read_bytes()
    backups_before = set(paths.backups.iterdir())
    with pytest.raises(module.BackupError, match="스키마"):
        service.restore(candidate, exclusive_access=True)
    assert db.path.read_bytes() == before
    assert set(paths.backups.iterdir()) == backups_before


def test_restore_accepts_sql_formatting_and_sqlite_statistics(prepared):
    module = backup_module()
    db, paths = prepared
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    candidate = service.backup()
    connection = sqlite3.connect(candidate)
    try:
        original = connection.execute("SELECT sql FROM sqlite_master WHERE name='report_supplemental_inputs_reject_locked_update'").fetchone()[0]
        connection.execute("DROP TRIGGER report_supplemental_inputs_reject_locked_update")
        connection.execute(original.replace("CREATE TRIGGER", "create /* formatting only */ trigger").replace("BEFORE UPDATE", "before\n\tupdate"))
        connection.execute("ANALYZE")
        connection.commit()
        assert connection.execute("SELECT 1 FROM sqlite_master WHERE name='sqlite_stat1'").fetchone()
    finally:
        connection.close()
    assert service.restore(candidate, exclusive_access=True).pre_restore_backup.is_file()
