from __future__ import annotations

import hashlib
from decimal import Decimal
from importlib import import_module

import pytest
from openpyxl import load_workbook

from app.db import Database
from app.repositories.masters import MasterRepository
from app.repositories.monthly_inputs import MonthLockedError, MonthlyInputRepository
from tests.fixtures.build_transport_fixture import (
    REGULAR_SHEET,
    SUBCONTRACT_SHEETS,
    build_transport_fixture,
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
            "SELECT destination_id, unresolved_alias FROM transport_entries "
            "WHERE destination_id IS NULL"
        ).fetchall()
        batch = connection.execute("SELECT * FROM import_batches").fetchone()
    assert [tuple(row) for row in unresolved] == [(None, "Unknown Plant")]
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
