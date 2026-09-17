from __future__ import annotations

import hashlib
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal
from importlib import import_module
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.db import Database
from app.repositories.masters import MasterRepository
from app.repositories.monthly_inputs import MonthLockedError, MonthlyInputRepository
from tests.fixtures.build_transport_fixture import (
    REGULAR_SHEET,
    SUBCONTRACT_SHEETS,
    build_structural_clone_fixture,
    build_transport_fixture,
    patch_formula_cached_values,
)


@pytest.fixture
def api():
    try:
        parser_module = import_module("app.importers.hwaseong_workbook")
        repository_module = import_module("app.repositories.transport_entries")
        service_module = import_module("app.services.import_service")
    except ModuleNotFoundError as error:
        pytest.fail(f"transport import API is not implemented: {error}")
    return parser_module, repository_module, service_module


@pytest.fixture
def fixture_path(tmp_path):
    return build_transport_fixture(tmp_path / "transport.xlsx")


@pytest.fixture
def database(tmp_path):
    result = Database(tmp_path / "app.db")
    result.migrate()
    masters = MasterRepository(result)
    destination = masters.create_destination("Known Destination", 1)
    masters.add_alias(destination.id, "Known Plant", "transport")
    return result


def test_parser_normalizes_day_columns_and_fractional_trips(api, fixture_path):
    parser_module, _, _ = api

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    matching = [
        row
        for row in rows
        if row.source_sheet == REGULAR_SHEET and row.day == 12
    ]
    assert len(matching) == 1
    assert matching[0].trip_count == Decimal("1.25")
    assert matching[0].cost_won == 2_500
    assert matching[0].vehicle_driver_group == "5톤 / 기사A"


def test_parser_skips_repeated_headers_and_displayed_totals(api, fixture_path):
    parser_module, _, _ = api

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5
    assert {row.source_sheet for row in rows} == {REGULAR_SHEET, *SUBCONTRACT_SHEETS}
    assert not {"운반지역", "합계"} & {row.destination_alias for row in rows}


def test_parser_accepts_anonymized_real_workbook_structure_and_formula_caches(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "structural-clone.xlsx")

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    assert len(rows) == 5
    assert any(
        row.source_sheet == REGULAR_SHEET
        and row.source_row == 3
        and row.trip_count == Decimal("1.25")
        and row.cost_won == 2_500
        for row in rows
    )
    assert any(
        row.source_sheet == REGULAR_SHEET and row.source_row == 5 for row in rows
    )
    assert not {"합계", "총계", "계"} & {row.destination_alias for row in rows}


def test_hidden_regular_detail_rows_are_parsed(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].row_dimensions[5].hidden = True
    workbook.save(fixture_path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert any(row.source_row == 5 for row in rows)


def test_hidden_subcontract_detail_rows_are_parsed(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[SUBCONTRACT_SHEETS[0]].row_dimensions[2].hidden = True
    workbook.save(fixture_path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert any(
        row.source_sheet == SUBCONTRACT_SHEETS[0] and row.source_row == 2
        for row in rows
    )


def test_parser_rejects_missing_required_sheet(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    del workbook[SUBCONTRACT_SHEETS[-1]]
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="required sheet"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_missing_required_header(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(2, 36, "not-unit-cost")
    workbook[REGULAR_SHEET].cell(4, 36, "not-unit-cost")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="header"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_requires_ah_total_count_header(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(2, 34, "not-total-count")
    workbook[REGULAR_SHEET].cell(4, 34, "not-total-count")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="header"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_regular_total_count_mismatch(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(3, 34, 2)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.TotalCountMismatchError, match="row 3"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("coordinate", "message"),
    [("AH3", "cached total count"), ("AK3", "cached subtotal")],
)
def test_parser_rejects_missing_regular_formula_cache(
    api, tmp_path, coordinate, message
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "missing-cache.xlsx")
    patch_formula_cached_values(path, 1, {coordinate: None})

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_missing_subcontract_formula_cache(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "missing-sub-cache.xlsx")
    patch_formula_cached_values(path, 2, {"I3": None})

    with pytest.raises(parser_module.WorkbookStructureError, match="amount"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_regular_sheet_month_mismatch(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "august.xlsx")

    with pytest.raises(parser_module.WorkbookStructureError, match="sheet month"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-09")


def test_parser_rejects_invalid_generated_calendar_date(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "february.xlsx", month=2)
    workbook = load_workbook(path)
    sheet = workbook["화성운반비내역(2월)"]
    sheet.cell(3, 33, 1)  # day 31
    workbook.save(path)
    workbook.close()
    patch_formula_cached_values(path, 1, {"AH3": "2.25", "AK3": "4500"})

    with pytest.raises(parser_module.WorkbookStructureError, match="calendar date"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-02")


def test_parser_rejects_data_row_missing_regular_destination(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(5, 2).value = None
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="destination"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_data_row_with_invalid_regular_unit_cost(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(5, 36, "invalid")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="unit cost"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_nontext_regular_destination(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(5, 2, 123)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="destination"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_data_row_missing_subcontract_destination(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 2).value = None
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="destination"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("column", "value", "message"),
    [(1, None, "date"), (9, None, "amount"), (9, "invalid", "amount")],
)
def test_parser_rejects_missing_or_invalid_subcontract_fields(
    api, fixture_path, column, value, message
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, column).value = value
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_nontext_subcontract_destination(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 2, 123)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="destination"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_subtotal_mismatch(api, tmp_path):
    parser_module, _, _ = api
    path = build_transport_fixture(
        tmp_path / "subtotal-mismatch.xlsx", subtotal_mismatch=True
    )

    with pytest.raises(parser_module.SubtotalMismatchError, match="row 3"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_import_is_idempotent_and_stores_file_hash(api, database, fixture_path):
    _, _, service_module = api
    service = service_module.TransportImportService(database)

    first = service.import_transport(fixture_path, "2026-08")
    second = service.import_transport(fixture_path, "2026-08")

    assert first.inserted_count == 5
    assert second.status == "duplicate"
    assert second.inserted_count == 0
    with database.connection() as connection:
        batches = connection.execute("SELECT * FROM import_batches").fetchall()
        entries = connection.execute("SELECT * FROM transport_entries").fetchall()
    assert len(batches) == 1
    assert len(entries) == 5
    assert batches[0]["file_sha256"] == hashlib.sha256(
        fixture_path.read_bytes()
    ).hexdigest()


def test_import_snapshots_once_then_hashes_and_parses_the_same_bytes(
    api, database, fixture_path
):
    parser_module, repository_module, service_module = api
    original_bytes = fixture_path.read_bytes()
    original_hash = hashlib.sha256(original_bytes).hexdigest()

    class MutatingParser:
        snapshot_path: Path | None = None
        snapshot_bytes: bytes | None = None

        def parse(self, path, report_month):
            self.snapshot_path = Path(path)
            self.snapshot_bytes = self.snapshot_path.read_bytes()
            fixture_path.write_bytes(b"source changed after snapshot")
            return []

    class RecordingRepository:
        source_filename: str | None = None
        file_sha256: str | None = None

        def import_entries(self, **kwargs):
            self.source_filename = kwargs["source_filename"]
            self.file_sha256 = kwargs["file_sha256"]
            return repository_module.ImportCommitResult(1, "imported", 0, ())

    parser = MutatingParser()
    repository = RecordingRepository()
    result = service_module.TransportImportService(
        database, parser=parser, repository=repository
    ).import_transport(fixture_path, "2026-08")

    assert parser.snapshot_bytes == original_bytes
    assert parser.snapshot_path is not None
    assert parser.snapshot_path != fixture_path
    assert not parser.snapshot_path.exists()
    assert result.file_sha256 == original_hash
    assert repository.file_sha256 == original_hash
    assert repository.source_filename == "transport.xlsx"


@pytest.mark.parametrize(
    ("limit_name", "limit", "message"),
    [
        ("MAX_INPUT_BYTES", 1, "input byte limit"),
        ("MAX_ZIP_MEMBERS", 1, "member count"),
        ("MAX_ZIP_MEMBER_BYTES", 1, "member size"),
        ("MAX_ZIP_UNCOMPRESSED_BYTES", 1, "uncompressed size"),
    ],
)
def test_import_rejects_oversized_or_zip_bomb_like_inputs(
    api, database, fixture_path, monkeypatch, limit_name, limit, message
):
    _, _, service_module = api
    monkeypatch.setattr(service_module, limit_name, limit)

    with pytest.raises(service_module.ImportFileLimitError, match=message):
        service_module.TransportImportService(database).import_transport(
            fixture_path, "2026-08"
        )

    with database.connection() as connection:
        assert (
            connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0]
            == 0
        )


def test_snapshot_temp_file_is_removed_when_parser_fails(api, database, fixture_path):
    parser_module, _, service_module = api

    class FailingParser:
        snapshot_path: Path | None = None

        def parse(self, path, report_month):
            self.snapshot_path = Path(path)
            raise parser_module.WorkbookStructureError("forced parse failure")

    parser = FailingParser()
    service = service_module.TransportImportService(database, parser=parser)

    with pytest.raises(parser_module.WorkbookStructureError, match="forced"):
        service.import_transport(fixture_path, "2026-08")

    assert parser.snapshot_path is not None
    assert not parser.snapshot_path.exists()


def test_concurrent_identical_imports_commit_one_batch(api, database, fixture_path):
    _, _, service_module = api
    service = service_module.TransportImportService(database)

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = [
            future.result()
            for future in (
                executor.submit(service.import_transport, fixture_path, "2026-08"),
                executor.submit(service.import_transport, fixture_path, "2026-08"),
            )
        ]

    assert sorted(result.status for result in results) == [
        "duplicate",
        "imported_with_errors",
    ]
    assert sorted(result.inserted_count for result in results) == [0, 5]
    with database.connection() as connection:
        assert (
            connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0]
            == 1
        )
        assert (
            connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0]
            == 5
        )


def test_unknown_alias_is_imported_as_blocking_unresolved_entry(
    api, database, fixture_path
):
    _, _, service_module = api

    result = service_module.TransportImportService(database).import_transport(
        fixture_path, "2026-08"
    )

    assert result.status == "imported_with_errors"
    assert result.blocking_errors == ("Unknown destination alias: Unknown Plant",)
    with database.connection() as connection:
        unresolved = connection.execute(
            "SELECT destination_id, unresolved_alias, source_alias "
            "FROM transport_entries "
            "WHERE destination_id IS NULL"
        ).fetchall()
        resolved_aliases = connection.execute(
            "SELECT DISTINCT source_alias FROM transport_entries "
            "WHERE destination_id IS NOT NULL"
        ).fetchall()
        batch = connection.execute("SELECT * FROM import_batches").fetchone()
    assert [tuple(row) for row in unresolved] == [
        (None, "Unknown Plant", "Unknown Plant")
    ]
    assert [row["source_alias"] for row in resolved_aliases] == ["Known Plant"]
    assert "Unknown Plant" in batch["error_summary"]


def test_structural_failure_is_atomic(api, database, tmp_path):
    parser_module, _, service_module = api
    path = build_transport_fixture(
        tmp_path / "subtotal-mismatch.xlsx", subtotal_mismatch=True
    )

    with pytest.raises(parser_module.SubtotalMismatchError):
        service_module.TransportImportService(database).import_transport(
            path, "2026-08"
        )

    with database.connection() as connection:
        assert connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0] == 0
        assert connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0] == 0


def test_locked_month_rejects_import_without_partial_batch(
    api, database, fixture_path
):
    _, _, service_module = api
    MonthlyInputRepository(database).finalize_month("2026-08", "revision-1")

    with pytest.raises(MonthLockedError, match="2026-08"):
        service_module.TransportImportService(database).import_transport(
            fixture_path, "2026-08"
        )

    with database.connection() as connection:
        assert connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0] == 0
        assert connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0] == 0
