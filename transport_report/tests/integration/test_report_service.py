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


def test_nonregular_is_report_cost_once_and_never_destination_cost(prepared):
    from decimal import Decimal
    from openpyxl import load_workbook
    from pptx import Presentation
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("UPDATE transport_entries SET cost_won=2000000")
        c.execute("UPDATE transport_entries SET cost_won=2475000 WHERE destination_id=1")
        c.execute("INSERT INTO transport_entries(import_batch_id,report_month,destination_id,source_sheet,source_row,transport_day,transport_type,trip_count_text,cost_won,source_alias) VALUES (1,'2026-08',1,'용차',90,1,'nonregular','1',4510000,'원천01')")
        c.execute("UPDATE report_supplemental_inputs SET nonregular_actual_quantity_text='123', nonregular_actual_cost_won=2880000 WHERE report_month='2026-08'")
        c.commit()
    service = module.ReportService(db, paths, clock=lambda: NOW)
    bundle, snapshot = service.prepare()
    assert sum(r.calculation.actual_cost_won for r in bundle.report.rows if r) == 24475000
    assert bundle.report.nonregular.calculation.actual_cost_won == 2880000
    assert bundle.report.total.calculation.actual_cost_won == 27355000
    assert bundle.report.total.calculation.actual_quantity == Decimal(60000)
    assert bundle.report.charts.actual_cost_won_by_month["2026-08"] == 27355000
    assert any(r["transport_type"] == "nonregular" and r["cost_won"] == 4510000 for r in json.loads(snapshot)["transport_entries"])
    assert all("용차!90" not in o.source_locator for o in bundle.operations if o.record_kind == "destination")
    result = service.generate()
    book = load_workbook(result.directory / "review.xlsx", data_only=True)
    try:
        assert any(cell.value == 27355000 for row in book["월간 종합"] for cell in row)
    finally:
        book.close()
    ppt = Presentation(result.directory / "report.pptx")
    table = next(s.table for s in ppt.slides[1].shapes if s.name == "report.monthly_table")
    assert table.cell(17, 6).text == "27,355"
    assert table.cell(16, 5).text == "2,880"
    from app.main import create_app
    from tests.integration.test_web_routes import WebClient
    page = WebClient(create_app(database=db, paths=paths)).get("/months/2026-08/preview").text
    assert "27355000" in page and "28985000" not in page


def test_nonregular_missing_pair_blocks_generation(prepared):
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("UPDATE report_supplemental_inputs SET nonregular_actual_quantity_text=NULL, nonregular_actual_cost_won=NULL WHERE report_month='2026-08'")
        c.commit()
    with pytest.raises(module.ReportGenerationError) as error:
        module.ReportService(db, paths).generate()
    assert "MISSING_REPORT_SUPPLEMENT" in {i.code for i in error.value.issues}
    assert not list(paths.outputs.iterdir())


def test_backup_without_hardlink_support_is_atomic_and_collision_safe(prepared, monkeypatch):
    module = backup_module()
    db, paths = prepared
    monkeypatch.setattr(module.os, "link", lambda *args: (_ for _ in ()).throw(OSError(95, "hard links unsupported")))
    service = module.BackupService(db, paths.backups, clock=lambda: NOW)
    first = service.backup()
    original = first.read_bytes()
    second = service.backup()
    assert first != second
    assert first.read_bytes() == original
    assert service._validate(second) == 6
    assert set(paths.backups.iterdir()) == {first, second}


class SimulatedTermination(BaseException):
    pass


def test_canonical_total_optional_nonregular_cost_preserves_quantity_and_missing_pair():
    from decimal import Decimal
    from app.domain.calculations import calculate_destination, calculate_total
    from app.domain.models import Destination
    rules = [Destination(1, "가상", 1, True, True, None, True, True, True)]
    direct = {1: calculate_destination(Decimal(10), 100, Decimal(20), 200, destination_id=1)}
    extra = calculate_destination(Decimal(999), 30, Decimal(888), 50)
    total = calculate_total(direct, rules, nonregular=extra)
    assert (total.planned_quantity, total.actual_quantity) == (Decimal(10), Decimal(20))
    assert (total.planned_cost_won, total.actual_cost_won) == (130, 250)
    assert calculate_total(direct, rules).actual_cost_won == 200
    missing = calculate_destination(None, None, None, None)
    assert calculate_total(direct, rules, nonregular=missing).actual_cost_won is None


@pytest.mark.parametrize("boundary", ["render", "before_publish", "after_publish"])
def test_terminated_run_has_durable_pending_state_and_recovers(prepared, monkeypatch, boundary):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    original_rename = service._rename
    def interrupted_render(*args):
        raise SimulatedTermination()
    def interrupted_publish(*args):
        if boundary == "after_publish":
            original_rename(*args)
        raise SimulatedTermination()
    monkeypatch.setattr(service, "_charts" if boundary == "render" else "_rename", interrupted_render if boundary == "render" else interrupted_publish)
    with pytest.raises(SimulatedTermination):
        service.generate()
    with db.connection() as c:
        rows = c.execute("SELECT * FROM report_runs").fetchall()
    assert len(rows) == 1 and rows[0]["status"] == "pending"
    module.ReportService(db, paths).recover_pending()
    with db.connection() as c:
        rows = c.execute("SELECT * FROM report_runs").fetchall()
    if boundary == "after_publish":
        assert len(rows) == 1 and rows[0]["status"] == "completed"
        assert len(list(paths.outputs.rglob("*.png"))) == 3
    else:
        assert rows == []
        assert not list(paths.outputs.rglob("*.*"))


def test_application_start_recovers_interrupted_publication(prepared, monkeypatch):
    from app.main import create_app
    from tests.integration.test_web_routes import WebClient
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    monkeypatch.setattr(service, "_audit", lambda *args: (_ for _ in ()).throw(SimulatedTermination()))
    with pytest.raises(SimulatedTermination):
        service.generate()
    assert WebClient(create_app(database=db, paths=paths)).get("/").status_code == 200
    with db.connection() as c:
        assert c.execute("SELECT status FROM report_runs").fetchone()[0] == "completed"


def test_pending_recovery_does_not_interrupt_live_generation(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    def during_render(bundle, directory):
        with pytest.raises(module.ReportGenerationError) as error:
            module.ReportService(db, paths).recover_pending()
        assert error.value.issues[0].code == "REPORT_GENERATION_BUSY"
        assert directory.exists()
        raise OSError("finish injected test")
    monkeypatch.setattr(service, "_charts", during_render)
    with pytest.raises(module.ReportGenerationError):
        service.generate()
    with db.connection() as c:
        assert not c.execute("SELECT * FROM report_runs").fetchall()


def test_pending_recovery_removes_damaged_published_run(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    monkeypatch.setattr(service, "_audit", lambda *args: (_ for _ in ()).throw(SimulatedTermination()))
    with pytest.raises(SimulatedTermination):
        service.generate()
    next(paths.outputs.rglob("*.png")).write_bytes(b"truncated")
    module.ReportService(db, paths).recover_pending()
    with db.connection() as c:
        assert not c.execute("SELECT * FROM report_runs").fetchall()
    assert not list(paths.outputs.rglob("*.*"))


def test_nonregular_source_is_audit_evidence_without_destination_alias_gate(prepared):
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("INSERT INTO transport_entries(import_batch_id,report_month,unresolved_alias,source_sheet,source_row,transport_day,transport_type,trip_count_text,cost_won,source_alias) VALUES (1,'2026-08','용차 원천','용차',90,1,'nonregular','1',4510000,'용차 원천')")
        c.commit()
    bundle, snapshot = module.ReportService(db, paths).prepare()
    assert bundle.validation.can_generate
    assert "용차 원천" in snapshot
    nonregular = next(o for o in bundle.operations if o.record_kind == "nonregular" and o.value_role == "actual")
    assert "4,510,000" in nonregular.source_locator and "보고 합산 제외" in nonregular.source_locator


def test_pending_cleanup_database_failure_remains_recoverable(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    monkeypatch.setattr(service, "_charts", lambda *args: (_ for _ in ()).throw(OSError("render failure")))
    monkeypatch.setattr(service, "_discard_pending", lambda *args: (_ for _ in ()).throw(sqlite3.OperationalError("D:/private database busy")))
    with pytest.raises(module.ReportGenerationError) as error:
        service.generate()
    assert "D:/private" not in str(error.value)
    with db.connection() as c:
        assert c.execute("SELECT status FROM report_runs").fetchone()[0] == "pending"
    module.ReportService(db, paths).recover_pending()
    with db.connection() as c:
        assert not c.execute("SELECT * FROM report_runs").fetchall()


def test_recovery_preserves_foreign_collision_directory_before_publish(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    target = paths.outputs / "2026-08" / "run-20260902-101112"
    def collision_before_publish(*args):
        target.mkdir()
        (target / "existing.txt").write_text("keep", encoding="utf-8")
        raise SimulatedTermination()
    monkeypatch.setattr(service, "_rename", collision_before_publish)
    with pytest.raises(SimulatedTermination):
        service.generate()
    module.ReportService(db, paths).recover_pending()
    assert (target / "existing.txt").read_text(encoding="utf-8") == "keep"
    with db.connection() as c:
        assert not c.execute("SELECT * FROM report_runs").fetchall()


def test_termination_before_pending_commit_creates_no_staging_directory(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths)
    monkeypatch.setattr(service, "_start_pending", lambda *args: (_ for _ in ()).throw(SimulatedTermination()))
    with pytest.raises(SimulatedTermination):
        service.generate()
    assert not list(paths.outputs.rglob(".run-*"))


def test_recovery_before_staging_creation_preserves_foreign_directory(prepared, monkeypatch):
    module = service_module()
    db, paths = prepared
    service = module.ReportService(db, paths, clock=lambda: NOW)
    original = service._start_pending
    target = paths.outputs / "2026-08" / "run-20260902-101112"
    def interrupted_start(*args):
        original(*args)
        target.mkdir()
        (target / "existing.txt").write_text("keep", encoding="utf-8")
        raise SimulatedTermination()
    monkeypatch.setattr(service, "_start_pending", interrupted_start)
    with pytest.raises(SimulatedTermination):
        service.generate()
    module.ReportService(db, paths).recover_pending()
    assert (target / "existing.txt").read_text(encoding="utf-8") == "keep"


def test_legacy_cost_history_does_not_require_new_historical_supplemental_inputs(prepared):
    module = service_module()
    db, paths = prepared
    with db.connection() as c:
        c.execute("INSERT INTO destinations(id,name,display_order,active,required_for_report,include_quantity_total) VALUES (13,'과거 비정규',13,0,0,0)")
        c.execute("UPDATE monthly_actual_costs SET source_type='legacy_workbook' WHERE report_month<'2026-08'")
        months = [r[0] for r in c.execute("SELECT DISTINCT report_month FROM monthly_actual_costs WHERE report_month<'2026-08'")]
        for month in months:
            c.execute("INSERT INTO monthly_actual_costs(report_month,destination_id,cost_won,source_type) VALUES (?,13,300000,'legacy_workbook')", (month,))
            c.execute("INSERT INTO monthly_plans(report_month,destination_id,quantity_ea_text,cost_won) VALUES (?,13,'0',200000)", (month,))
        c.execute("DELETE FROM report_supplemental_inputs WHERE report_month<'2026-08'")
        c.commit()
    bundle, _ = module.ReportService(db, paths).prepare()
    assert bundle.validation.can_generate
    assert bundle.report.charts.actual_cost_won_by_month["2026-07"] == 25500000
    assert bundle.report.charts.planned_cost_won_by_month["2026-07"] == 24200000
    assert bundle.report.charts.actual_cost_won_by_month["2026-08"] == 25200000
