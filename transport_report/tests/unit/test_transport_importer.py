from __future__ import annotations

import hashlib
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from decimal import Decimal
from importlib import import_module
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as WorksheetImage
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table
from PIL import Image as PillowImage

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


def _parsed_entry(parser_module, **overrides):
    values = {
        "report_month": "2026-08",
        "destination_alias": "Known Plant",
        "source_sheet": REGULAR_SHEET,
        "source_row": 3,
        "source_date": date(2026, 8, 1),
        "day": 1,
        "transport_type": "regular",
        "trip_count": Decimal("1"),
        "unit_rate_won": 1,
        "cost_won": 1,
    }
    values.update(overrides)
    return parser_module.ParsedTransportEntry(**values)


def _rewrite_zip_member(path: Path, member_name: str, transform) -> None:
    with ZipFile(path, "r") as source:
        members = [
            (info, transform(source.read(info.filename)))
            if info.filename == member_name
            else (info, source.read(info.filename))
            for info in source.infolist()
        ]

    temporary = path.with_name(path.name + ".rewritten")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as destination:
        for info, data in members:
            destination.writestr(info, data)
    temporary.replace(path)


def _fail_if_openpyxl_loads(*args, **kwargs):
    pytest.fail("load_workbook was called before workbook resource preflight")


def _add_chartsheet(path: Path) -> None:
    workbook = load_workbook(path)
    source = workbook[REGULAR_SHEET]
    chart = BarChart()
    chart.add_data(Reference(source, min_col=1, min_row=1, max_row=2))
    workbook.create_chartsheet("Chart").add_chart(chart)
    workbook.save(path)
    workbook.close()


def _add_shared_strings(path: Path, records: bytes) -> None:
    _add_zip_member(
        path,
        "xl/sharedStrings.xml",
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + records
        + b"</sst>",
    )
    _rewrite_zip_member(
        path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/sharedStrings.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            b"</Types>",
        ),
    )


def _add_comment(path: Path, reference: str = "A1") -> None:
    workbook = load_workbook(path)
    workbook[REGULAR_SHEET][reference].comment = Comment("note", "author")
    workbook.save(path)
    workbook.close()


def _remove_zip_member(path: Path, member_name: str) -> None:
    with ZipFile(path, "r") as source:
        members = [
            (info, source.read(info.filename))
            for info in source.infolist()
            if info.filename != member_name
        ]

    temporary = path.with_name(path.name + ".rewritten")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as destination:
        for info, data in members:
            destination.writestr(info, data)
    temporary.replace(path)


def _add_zip_member(path: Path, member_name: str, data: bytes) -> None:
    with ZipFile(path, "a", compression=ZIP_DEFLATED) as archive:
        archive.writestr(member_name, data)


def _rename_zip_member(
    path: Path, old_member_name: str, new_member_name: str, transform=lambda data: data
) -> None:
    with ZipFile(path, "r") as source:
        members = [
            (
                new_member_name if info.filename == old_member_name else info.filename,
                transform(source.read(info.filename))
                if info.filename == old_member_name
                else source.read(info.filename),
            )
            for info in source.infolist()
        ]

    temporary = path.with_name(path.name + ".renamed")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as destination:
        for member_name, data in members:
            destination.writestr(member_name, data)
    temporary.replace(path)


def _add_workbook_pivot_cache_reference(path: Path, target: str) -> None:
    _rewrite_zip_member(
        path,
        "xl/workbook.xml",
        lambda data: data.replace(
            b"<calcPr ",
            b'<pivotCaches xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b'<pivotCache cacheId="1" r:id="rId99"/>'
            b"</pivotCaches><calcPr ",
            1,
        ),
    )
    _rewrite_zip_member(
        path,
        "xl/_rels/workbook.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="rId99" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
            + f'Target="{target}"/>'.encode()
            + b"</Relationships>",
            1,
        ),
    )


def _add_worksheet_chart(path: Path) -> None:
    workbook = load_workbook(path)
    source = workbook[REGULAR_SHEET]
    chart = BarChart()
    chart.add_data(Reference(source, min_col=1, min_row=1, max_row=2))
    source.add_chart(chart, "A10")
    workbook.save(path)
    workbook.close()


def _add_worksheet_hyperlink(
    path: Path,
    *,
    sheet_name: str = REGULAR_SHEET,
    target: str | None = "https://example.test/report",
    location: str | None = None,
) -> None:
    workbook = load_workbook(path)
    cell = workbook[sheet_name]["A1"]
    if target is None:
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location)
    else:
        cell.hyperlink = target
    workbook.save(path)
    workbook.close()


def _external_hyperlink_relationships(count: int, *, start: int = 1) -> bytes:
    return b"".join(
        b'<Relationship Id="external-%d" ' % relationship_id
        + b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        + b'Target="https://example.test/%d" TargetMode="External"/>'
        % relationship_id
        for relationship_id in range(start, start + count)
    )


def _relationship_record_count(path: Path) -> int:
    with ZipFile(path) as archive:
        return sum(
            archive.read(info.filename).count(b"<Relationship ")
            for info in archive.infolist()
            if info.filename.endswith(".rels")
        )


def _add_drawing_hyperlink(path: Path) -> None:
    _add_worksheet_chart(path)
    _rewrite_zip_member(
        path,
        "xl/drawings/drawing1.xml",
        lambda data: data.replace(
            b'<cNvPr id="1" name="Chart 1"/>',
            b'<cNvPr id="1" name="Chart 1">'
            b'<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            b'r:id="rId2"/></cNvPr>',
            1,
        ),
    )
    _rewrite_zip_member(
        path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="rId2" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            b'Target="https://example.test/chart" TargetMode="External"/>'
            b"</Relationships>",
            1,
        ),
    )


def _add_worksheet_table(path: Path) -> None:
    workbook = load_workbook(path)
    source = workbook[REGULAR_SHEET]
    source["AZ1"] = "First"
    source["BA1"] = "Second"
    source["AZ2"] = 1
    source["BA2"] = 2
    source.add_table(Table(displayName="AuditTable", ref="AZ1:BA2"))
    workbook.save(path)
    workbook.close()


def _add_worksheet_chart_and_image(path: Path) -> None:
    image_bytes = BytesIO()
    PillowImage.new("RGB", (2, 2), "red").save(image_bytes, format="PNG")
    image_bytes.seek(0)
    image = PillowImage.open(image_bytes)
    workbook = load_workbook(path)
    source = workbook[REGULAR_SHEET]
    chart = BarChart()
    chart.add_data(Reference(source, min_col=1, min_row=1, max_row=2))
    source.add_chart(chart, "A10")
    source.add_image(WorksheetImage(image), "F10")
    workbook.save(path)
    workbook.close()


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


def test_parser_rejects_repeated_regular_header_with_uncached_ancillary_formula(
    api, fixture_path
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(4, 35, "=1")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match=r"row 4"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_initial_regular_header_with_uncached_ancillary_formula(
    api, fixture_path
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(2, 35, "=1")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="header"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


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


def test_parser_skips_anonymized_regular_footer_without_destination(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "regular-footer.xlsx", include_regular_footer=True
    )

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    assert not any(
        row.source_sheet == REGULAR_SHEET and row.source_row == 113 for row in rows
    )


def test_parser_reports_uncached_regular_footer_al_formula_at_cell(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "regular-footer-uncached-al.xlsx", include_regular_footer=True
    )
    patch_formula_cached_values(path, 1, {"AL113": None})

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 113 cell AL113 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize(("column", "coordinate"), [(34, "AH113"), (36, "AJ113")])
def test_parser_rejects_uncached_regular_footer_formula(
    api, tmp_path, column, coordinate
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / f"regular-footer-{coordinate}-formula.xlsx",
        include_regular_footer=True,
        regular_footer_uncached_formula_column=column,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"row 113 cell {coordinate} cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize("ancillary_value", ["=1", 1], ids=("formula", "value"))
def test_parser_rejects_regular_footer_with_ancillary_content(
    api, tmp_path, ancillary_value
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / f"regular-footer-ancillary-{ancillary_value!s}.xlsx",
        include_regular_footer=True,
        regular_footer_ancillary_value=ancillary_value,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row 113.*destination"
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_ak_only_regular_footer_lookalike(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "regular-footer-lookalike.xlsx",
        include_regular_footer_lookalike=True,
    )

    with pytest.raises(parser_module.WorkbookStructureError, match="destination"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_skips_both_inactive_regular_row_shapes_and_updates_group(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "inactive-rows.xlsx", include_inactive_regular_rows=True
    )

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    aliases = {row.destination_alias for row in rows}
    assert "Dormant Blank Destination" not in aliases
    assert "Dormant Zero Destination" not in aliases
    trailing_row = next(
        row
        for row in rows
        if row.source_sheet == REGULAR_SHEET and row.source_row == 8
    )
    assert trailing_row.vehicle_driver_group == "8-ton / Driver B"


def test_parser_rejects_missing_aj_cache_on_inactive_looking_row(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "inactive-missing-aj-cache.xlsx",
        include_inactive_regular_rows=True,
        missing_inactive_regular_unit_formula_cache=True,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 6 cell AJ6 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_missing_ah_cache_on_inactive_looking_row(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "inactive-missing-ah-cache.xlsx",
        include_inactive_regular_rows=True,
        missing_inactive_regular_count_formula_cache=True,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 6 cell AH6 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_missing_ak_cache_on_inactive_looking_row(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "inactive-missing-ak-cache.xlsx",
        include_inactive_regular_rows=True,
        missing_inactive_regular_subtotal_formula_cache=True,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 6 cell AK6 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_inactive_regular_row_with_uncached_ancillary_formula(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "inactive-ancillary-formula.xlsx",
        include_inactive_regular_rows=True,
        inactive_regular_ancillary_value="=1",
    )

    with pytest.raises(parser_module.WorkbookStructureError, match=r"row 6.*unit"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_blank_regular_row_with_uncached_ancillary_formula(
    api, fixture_path
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(6, 35, "=1")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row 6.*destination"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_numeric_string_regular_unit_rate(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(3, 36, "2000")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="unit cost"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("column", "value"),
    [
        (3, 1),
        (3, "malformed"),
        (34, 1),
        (34, "malformed"),
        (37, 1),
        (37, "malformed"),
        (36, "malformed"),
    ],
)
def test_parser_rejects_inactive_looking_row_with_invalid_activity_or_unit(
    api, fixture_path, column, value
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[REGULAR_SHEET]
    for day_column in range(3, 34):
        sheet.cell(5, day_column).value = None
    sheet.cell(5, 34, 0)
    sheet.cell(5, 36).value = None
    sheet.cell(5, 37, 0)
    sheet.cell(5, column, value)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


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
    [("AH3", "AH3"), ("AK3", "AK3")],
)
def test_parser_rejects_missing_regular_formula_cache(
    api, tmp_path, coordinate, message
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "missing-cache.xlsx")
    patch_formula_cached_values(path, 1, {coordinate: None})

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"row 3 cell {message} cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize(
    ("fixture_option", "coordinate"),
    [
        ("missing_regular_day_formula_cache", "N3"),
        ("missing_regular_unit_formula_cache", "AJ3"),
    ],
)
def test_parser_reports_missing_daily_and_unit_formula_caches(
    api, tmp_path, fixture_option, coordinate
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / f"missing-{coordinate}-cache.xlsx",
        **{fixture_option: True},
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"row 3 cell {coordinate} cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_missing_subcontract_formula_cache(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "missing-sub-cache.xlsx")
    patch_formula_cached_values(path, 2, {"I3": None})

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 3 cell I3 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_missing_subcontract_total_formula_cache(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "missing-total-cache.xlsx")
    patch_formula_cached_values(path, 2, {"I4": None})

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"row 4 cell I4 cached value is missing.*recalculate.*Excel",
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_stale_subcontract_total_formula_cache(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "stale-total-cache.xlsx")
    patch_formula_cached_values(path, 2, {"I4": "50001"})

    with pytest.raises(parser_module.SubcontractTotalMismatchError, match="row 4"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize("sheet_name", SUBCONTRACT_SHEETS)
def test_parser_requires_one_displayed_subcontract_total(api, fixture_path, sheet_name):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[sheet_name]
    for cell in sheet[3]:
        cell.value = None
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match="displayed cached total"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize("sheet_name", SUBCONTRACT_SHEETS)
def test_parser_rejects_non_numeric_subcontract_total(api, fixture_path, sheet_name):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[sheet_name].cell(3, 9, "not-a-total")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="cached total"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize("sheet_name", SUBCONTRACT_SHEETS)
def test_parser_rejects_mismatched_subcontract_total(api, fixture_path, sheet_name):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[sheet_name]
    sheet.cell(3, 9, sheet.cell(3, 9).value + 1)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.SubcontractTotalMismatchError, match="row 3"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_reconciles_anonymized_total_rows_in_columns_a_and_g(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "total-placements.xlsx")

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    assert len([row for row in rows if row.transport_type == "nonregular"]) == 3


def test_parser_rejects_total_row_with_uncached_ancillary_formula(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "uncached-total-ancillary-formula.xlsx",
        subcontract_total_ancillary_value="=1",
    )

    with pytest.raises(parser_module.WorkbookStructureError, match=r"row 4.*date"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize(
    ("value", "cached_value"),
    [(1, None), ("=1", "1")],
    ids=("cached-value", "cached-formula"),
)
def test_parser_rejects_total_row_with_ancillary_cached_value_or_formula(
    api, tmp_path, value, cached_value
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / f"total-ancillary-{cached_value or 'value'}.xlsx",
        subcontract_total_ancillary_value=value,
        subcontract_total_ancillary_cached_value=cached_value,
    )

    with pytest.raises(parser_module.WorkbookStructureError, match=r"row 4.*date"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_total_label_text_in_ancillary_detail_column(
    api, fixture_path
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[SUBCONTRACT_SHEETS[0]]
    sheet.insert_rows(3)
    sheet.cell(3, 3, "계")
    sheet.cell(3, 9, 50_000)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="date"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_subcontract_detail_after_displayed_total(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[SUBCONTRACT_SHEETS[0]]
    sheet.cell(4, 1, 2)
    sheet.cell(4, 2, date(2026, 8, 6))
    sheet.cell(4, 6, "Known Plant")
    sheet.cell(4, 8, "5톤")
    sheet.cell(4, 9, 60_000)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row 4.*detail after.*total"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_second_subcontract_displayed_total(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[SUBCONTRACT_SHEETS[0]]
    sheet.cell(4, 1, "총계")
    sheet.cell(4, 9, 50_000)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row 4.*second displayed total"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_accepts_actual_post_total_blank_template_and_formula_footer_shapes(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "post-total-footers.xlsx",
        include_subcontract_post_total_rows=True,
    )

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    assert len(rows) == 5


def test_parser_rejects_subcontract_header_with_uncached_ancillary_formula(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "subcontract-header-ancillary-formula.xlsx",
        subcontract_header_ancillary_value="=1",
    )

    with pytest.raises(parser_module.WorkbookStructureError, match="header"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


@pytest.mark.parametrize("sheet_name", [REGULAR_SHEET, *SUBCONTRACT_SHEETS[:1]])
def test_parser_rejects_worksheet_row_dimension_abuse(
    api, fixture_path, sheet_name
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[sheet_name].cell(2_001, 1).number_format = "0"
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row limit.*2000"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize("sheet_name", [REGULAR_SHEET, *SUBCONTRACT_SHEETS[:1]])
def test_parser_rejects_worksheet_column_dimension_abuse(
    api, fixture_path, sheet_name
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[sheet_name].cell(1, 129).number_format = "0"
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"column limit.*128"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_oversized_relevant_sheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_CELLS", 82)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"cell-record limit.*82"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_oversized_irrelevant_sheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    irrelevant = workbook.create_sheet("Irrelevant")
    for row_number in range(1, 85):
        irrelevant.cell(row_number, 1, row_number)
    workbook.save(fixture_path)
    workbook.close()
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_CELLS", 83)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Irrelevant.*cell-record limit.*83",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_large_total_cell_count_with_low_dimensions(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_TOTAL_CELLS", 100)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"workbook cell-record limit.*100",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("row_number", "column_number", "message"),
    [
        (20_001, 1, r"preflight row limit.*20000"),
        (1, 513, r"preflight column limit.*512"),
    ],
    ids=("row", "column"),
)
def test_preflight_rejects_extreme_references_before_openpyxl(
    api, fixture_path, monkeypatch, row_number, column_number, message
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(row_number, column_number, 1)
    workbook.save(fixture_path)
    workbook.close()
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_large_worksheet_xml_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_XML_BYTES", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"worksheet XML size limit.*1"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_huge_merged_range_on_irrelevant_sheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook.create_sheet("Irrelevant").cell(1, 1, 1)
    workbook.save(fixture_path)
    workbook.close()
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet5.xml",
        lambda data: data.replace(
            b"</worksheet>",
            b'<mergeCells count="1"><mergeCell ref="A1:SR20000"/>'
            b"</mergeCells></worksheet>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Irrelevant.*cell materialization limit.*50000",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_huge_hyperlink_range_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook.create_sheet("Irrelevant").cell(1, 1, 1)
    workbook.save(fixture_path)
    workbook.close()
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet5.xml",
        lambda data: data.replace(
            b"</worksheet>",
            b'<hyperlinks><hyperlink ref="A1:SR20000" location="x"/>'
            b"</hyperlinks></worksheet>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Irrelevant.*cell materialization limit.*50000",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_counts_rows_without_explicit_references_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook.create_sheet("Irrelevant").cell(1, 1, 1)
    workbook.save(fixture_path)
    workbook.close()
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet5.xml",
        lambda data: data.replace(
            b"</sheetData>", b'<row s="0"/>' * 20_000 + b"</sheetData>"
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"preflight row limit.*20000"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_duplicate_row_record_overflow_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook.create_sheet("Irrelevant").cell(1, 1, 1)
    workbook.save(fixture_path)
    workbook.close()
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet5.xml",
        lambda data: data.replace(
            b"</sheetData>", b'<row r="1"/>' * 20_000 + b"</sheetData>"
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row-record limit.*20000"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_column_dimension_record_overflow_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    definitions = b'<col min="1" max="1" width="10"/>' * 2_049
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet1.xml",
        lambda data: data.replace(
            b"<sheetData>",
            b"<cols>" + definitions + b"</cols><sheetData>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"column-dimension record limit.*2048",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_openpyxl_materializes_one_column_dimension_for_a_full_width_span(
    fixture_path,
):
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet1.xml",
        lambda data: data.replace(
            b"<sheetData>",
            b'<cols><col min="1" max="16384" width="10"/></cols><sheetData>',
        ),
    )

    workbook = load_workbook(fixture_path)
    dimensions = workbook[REGULAR_SHEET].column_dimensions
    assert len(dimensions) == 1
    assert (dimensions["A"].min, dimensions["A"].max) == (1, 16_384)
    workbook.close()


def test_preflight_accepts_full_width_style_column_span(api, fixture_path):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet1.xml",
        lambda data: data.replace(
            b"<sheetData>",
            b'<cols><col min="1" max="16384" style="1"/></cols><sheetData>',
        ),
    )

    parser_module._preflight_workbook_resources(fixture_path)


def test_preflight_accepts_actual_sized_column_dimension_record_shape(
    api, fixture_path
):
    parser_module, _, _ = api
    definitions = b"".join(
        f'<col min="{column}" max="{column}" width="10"/>'.encode()
        for column in range(1, 705)
    )
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet1.xml",
        lambda data: data.replace(
            b"<sheetData>", b"<cols>" + definitions + b"</cols><sheetData>"
        ),
    )

    parser_module._preflight_workbook_resources(fixture_path)


@pytest.mark.parametrize(
    "column_definition",
    [
        b'<col max="1"/>',
        b'<col min="1"/>',
        b'<col min="x" max="1"/>',
        b'<col min="2" max="1"/>',
        b'<col min="0" max="1"/>',
        b'<col min="1" max="16385"/>',
        b'<col min="' + b"9" * 10_000 + b'" max="1"/>',
    ],
    ids=(
        "missing-min",
        "missing-max",
        "noninteger",
        "reversed",
        "zero",
        "out-of-range",
        "pathological-integer",
    ),
)
def test_preflight_rejects_invalid_column_dimensions_before_openpyxl(
    api, fixture_path, monkeypatch, column_definition
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/sheet1.xml",
        lambda data: data.replace(
            b"<sheetData>", b"<cols>" + column_definition + b"</cols><sheetData>"
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"invalid column dimension",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_normal_styled_columns(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[REGULAR_SHEET]
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].hidden = True
    workbook.save(fixture_path)
    workbook.close()

    parser_module._preflight_workbook_resources(fixture_path)


def test_preflight_covers_nonstandard_sheet_relationship_openpyxl_would_load(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "xl/_rels/workbook.xml.rels",
        lambda data: data.replace(
            b"/relationships/worksheet\"",
            b"/relationships/dialogsheet\"",
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_CELLS", 82)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"cell-record limit.*82"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_uses_the_relationship_id_openpyxl_will_use(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "xl/workbook.xml",
        lambda data: data.replace(
            b"<sheet ",
            b'<sheet xmlns:x="urn:untrusted" x:id="rId2" ',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_CELLS", 82)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"cell-record limit.*82"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("member_name", "transform", "message"),
    [
        (
            "xl/_rels/workbook.xml.rels",
            lambda data: data.replace(
                b'Target="/xl/worksheets/sheet1.xml"',
                b'Target="../worksheets/sheet1.xml"',
                1,
            ),
            "worksheet relationship target",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            lambda data: data[:-1],
            "relationships XML",
        ),
        ("xl/worksheets/sheet1.xml", lambda data: data[:-1], "Worksheet XML"),
    ],
    ids=("unsafe-path", "malformed-relationships", "malformed-worksheet"),
)
def test_preflight_rejects_unsafe_or_malformed_xml_before_openpyxl(
    api, fixture_path, monkeypatch, member_name, transform, message
):
    parser_module, _, _ = api
    _rewrite_zip_member(fixture_path, member_name, transform)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_normal_synthetic_workbook(api, fixture_path):
    parser_module, _, _ = api

    parser_module._preflight_workbook_resources(fixture_path)


def test_preflight_accepts_valid_chartsheet_and_parser_reads_required_sheets(
    api, fixture_path
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_counts_chartsheets_in_workbook_sheet_limit(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEETS", 4)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"worksheet count limit.*4"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("target", "message"),
    [
        ("../chartsheets/sheet1.xml", "unsafe chartsheet relationship target"),
        ("/etc/passwd", "unsafe chartsheet relationship target"),
    ],
    ids=("traversal", "absolute-outside-xl"),
)
def test_preflight_rejects_unsafe_chartsheet_target_before_openpyxl(
    api, fixture_path, monkeypatch, target, message
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/_rels/workbook.xml.rels",
        lambda data: data.replace(
            b'Target="/xl/chartsheets/sheet1.xml"',
            f'Target="{target}"'.encode(),
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_external_chartsheet_target_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/_rels/workbook.xml.rels",
        lambda data: data.replace(
            b'Target="/xl/chartsheets/sheet1.xml"',
            b'Target="https://example.test/chart.xml" TargetMode="External"',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"external chartsheet relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_chartsheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    _remove_zip_member(fixture_path, "xl/chartsheets/sheet1.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Chartsheet XML is missing"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_chartsheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/chartsheets/sheet1.xml",
        lambda data: data[:-1],
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Chartsheet XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_oversized_chartsheet_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATED_XML_BYTES", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"chartsheet XML size limit.*1"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_worksheet_relationships_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data[:-1],
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Worksheet relationships XML is malformed",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_worksheet_hyperlink_relationship_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _remove_zip_member(fixture_path, "xl/worksheets/_rels/sheet1.xml.rels")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Worksheet hyperlink relationships are missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_wrong_type_worksheet_hyperlink_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b"/relationships/hyperlink\"",
            b"/relationships/image\"",
            1,
        ).replace(b' TargetMode="External"', b"", 1),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"invalid worksheet hyperlink relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    "relationship_type",
    [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink",
    ],
    ids=("transitional", "strict"),
)
def test_preflight_accepts_valid_external_worksheet_hyperlink(
    api, fixture_path, relationship_type
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            relationship_type.encode(),
            1,
        ),
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_accepts_location_only_worksheet_hyperlink(api, fixture_path):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path, target=None, location="Sheet2!A1")

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_accepts_internal_worksheet_hyperlink_relationship(
    api, fixture_path
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b'Target="https://example.test/report" TargetMode="External"',
            b'Target="../workbook.xml" TargetMode="Internal"',
            1,
        ),
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


@pytest.mark.parametrize(
    "target",
    ["", "   ", "https://example.test/&#x7f;report"],
    ids=("blank", "whitespace", "control-character"),
)
def test_preflight_rejects_invalid_external_hyperlink_target_before_openpyxl(
    api, fixture_path, monkeypatch, target
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b'Target="https://example.test/report"',
            f'Target="{target}"'.encode(),
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"invalid worksheet hyperlink relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_overlong_external_hyperlink_target_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_HYPERLINK_TARGET_LENGTH", 8)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"invalid worksheet hyperlink relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_uses_first_duplicate_worksheet_hyperlink_relationship(
    api, fixture_path
):
    parser_module, _, _ = api
    _add_worksheet_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            b'Target="/xl/workbook.xml"/></Relationships>',
            1,
        ),
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


@pytest.mark.parametrize(
    "target",
    ["../../../outside.xml", "../../xl/comments/comment1.xml"],
    ids=("outside-root", "escape-and-return"),
)
def test_preflight_rejects_unsafe_worksheet_relationship_target_before_openpyxl(
    api, fixture_path, monkeypatch, target
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b'Target="/xl/comments/comment1.xml"',
            f'Target="{target}"'.encode(),
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"unsafe worksheet relationship target",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_comment_part_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _remove_zip_member(fixture_path, "xl/comments/comment1.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Comment XML is missing"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_eager_drawing_part_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    source = workbook[REGULAR_SHEET]
    chart = BarChart()
    chart.add_data(Reference(source, min_col=1, min_row=1, max_row=2))
    source.add_chart(chart, "A10")
    workbook.save(fixture_path)
    workbook.close()
    _remove_zip_member(fixture_path, "xl/drawings/drawing1.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Worksheet related XML is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_chartsheet_drawing_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_chartsheet(fixture_path)
    _remove_zip_member(fixture_path, "xl/drawings/drawing1.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_worksheet_chart_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _remove_zip_member(fixture_path, "xl/charts/chart1.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_drawing_relationships_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _remove_zip_member(fixture_path, "xl/drawings/_rels/drawing1.xml.rels")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship references are missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_follows_table_part_id_regardless_of_relationship_type(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_table(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b"/relationships/table\"",
            b"/relationships/notTable\"",
            1,
        ),
    )
    _rewrite_zip_member(
        fixture_path,
        "xl/tables/table1.xml",
        lambda data: data[:-1],
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Related XML is malformed",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_external_hyperlink_used_as_table_part_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_table(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
        lambda data: data.replace(
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
            b'Target="/xl/tables/table1.xml"',
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            b'Target="https://example.test/table.xml" TargetMode="External"',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"unsupported external worksheet relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_unsafe_nested_relationship_target_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b'Target="/xl/charts/chart1.xml"',
            b'Target="../../../outside.xml"',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"unsafe relationship graph target",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_nested_xml_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/charts/chart1.xml",
        lambda data: data[:-1],
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Related XML is malformed",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_invalid_nested_relationship_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ',
            b"",
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph relationships are invalid",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_oversized_nested_xml_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    with ZipFile(fixture_path) as archive:
        protected_sizes = [
            archive.getinfo(member_name).file_size
            for member_name in (
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/drawings/drawing1.xml",
                "xl/charts/chart1.xml",
            )
        ]
    limit = max(protected_sizes) + 1
    _rewrite_zip_member(
        fixture_path,
        "xl/charts/chart1.xml",
        lambda data: data.replace(b"</chartSpace>", b" " * limit + b"</chartSpace>"),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATED_XML_BYTES", limit)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"related XML size limit.*{limit}",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_relationship_graph_cycle_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _add_zip_member(
        fixture_path,
        "xl/charts/_rels/chart1.xml.rels",
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="urn:test:cycle" '
        b'Target="/xl/drawings/drawing1.xml"/>'
        b"</Relationships>",
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship graph contains a cycle",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_counts_duplicate_relationship_edges_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="duplicate" Type="urn:test:duplicate" '
            b'Target="/xl/charts/chart1.xml"/></Relationships>',
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATIONSHIP_EDGES", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship edge limit.*1",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("relationships_member", "description"),
    (
        ("_rels/.rels", "Package relationships"),
        ("xl/_rels/workbook.xml.rels", "Workbook relationships"),
    ),
)
def test_preflight_rejects_too_many_package_relationship_records_before_openpyxl(
    api, fixture_path, monkeypatch, relationships_member, description
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        relationships_member,
        lambda data: data.replace(
            b"</Relationships>",
            _external_hyperlink_relationships(4_097) + b"</Relationships>",
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"{description} exceeds relationship record limit.*4096",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_too_many_unused_nested_relationship_records_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            _external_hyperlink_relationships(4_097) + b"</Relationships>",
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship graph relationships exceeds relationship record limit.*4096",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_total_relationship_records_across_many_parts_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    for sheet_name in (REGULAR_SHEET, *SUBCONTRACT_SHEETS):
        _add_worksheet_hyperlink(fixture_path, sheet_name=sheet_name)
    monkeypatch.setattr(
        parser_module,
        "MAX_PREFLIGHT_TOTAL_RELATIONSHIP_RECORDS",
        _relationship_record_count(fixture_path) - 1,
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"total relationship record limit",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_exact_relationship_record_part_limit_and_external_hyperlinks(
    api, fixture_path
):
    parser_module, _, _ = api
    relationship_limit = parser_module.MAX_PREFLIGHT_RELATIONSHIP_RECORDS
    _rewrite_zip_member(
        fixture_path,
        "_rels/.rels",
        lambda _: (
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            b'Target="xl/workbook.xml"/>'
            + _external_hyperlink_relationships(relationship_limit - 1)
            + b"</Relationships>"
        ),
    )
    _add_worksheet_hyperlink(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_rejects_relationship_part_limit_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATIONSHIP_PARTS", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship part limit.*1",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_relationship_depth_limit_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATIONSHIP_DEPTH", 0)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship depth limit.*0",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_relationship_queue_limit_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATIONSHIP_QUEUE", 0)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship queue limit.*0",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_bounds_relationship_roots_while_collecting_them(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    for sheet_name in (REGULAR_SHEET, SUBCONTRACT_SHEETS[0]):
        source = workbook[sheet_name]
        chart = BarChart()
        chart.add_data(Reference(source, min_col=1, min_row=1, max_row=2))
        source.add_chart(chart, "A10")
    workbook.save(fixture_path)
    workbook.close()
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATIONSHIP_QUEUE", 1)
    monkeypatch.setattr(
        parser_module,
        "_preflight_relationship_graph",
        lambda *args: pytest.fail("relationship roots were not bounded during collection"),
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"relationship queue limit.*1",
    ):
        parser_module._preflight_workbook_resources(fixture_path)


def test_preflight_rejects_missing_nested_image_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart_and_image(fixture_path)
    _remove_zip_member(fixture_path, "xl/media/image1.png")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_oversized_nested_binary_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart_and_image(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_RELATED_PART_BYTES", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"related part size limit.*1",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_nested_external_relationship_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b'Target="/xl/charts/chart1.xml"',
            b'Target="https://example.test/chart.xml" TargetMode="External"',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"unsupported external relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_external_hyperlink_used_as_drawing_part_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
            b'Target="/xl/charts/chart1.xml"',
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            b'Target="https://example.test/chart" TargetMode="External"',
            1,
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"unsupported external relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_valid_external_drawing_hyperlink(api, fixture_path):
    parser_module, _, _ = api
    _add_drawing_hyperlink(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_rejects_wrong_type_drawing_hyperlink_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_drawing_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b"/relationships/hyperlink\"",
            b"/relationships/image\"",
            1,
        ).replace(b' TargetMode="External"', b"", 1),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"invalid drawing hyperlink relationship",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_conflicting_duplicate_drawing_relationship(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_drawing_hyperlink(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="rId2" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            b'Target="/xl/charts/chart1.xml"/></Relationships>',
            1,
        ),
    )

    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"content type conflicts with image relationship semantics",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_workbook_pivot_cache_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_workbook_pivot_cache_reference(
        fixture_path, "/xl/pivotCache/missing.xml"
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_follows_pivot_cache_records_chain_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_workbook_pivot_cache_reference(
        fixture_path, "/xl/pivotCache/pivotCacheDefinition1.xml"
    )
    _add_zip_member(
        fixture_path,
        "xl/pivotCache/pivotCacheDefinition1.xml",
        b'<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        b'r:id="rId1"/>',
    )
    _add_zip_member(
        fixture_path,
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" '
        b'Target="/xl/pivotCache/missingRecords.xml"/></Relationships>',
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
            b"</Types>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_follows_table_query_chain_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_table(fixture_path)
    _add_zip_member(
        fixture_path,
        "xl/tables/_rels/table1.xml.rels",
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" '
        b'Target="/xl/queryTables/missing.xml"/></Relationships>',
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_valid_worksheet_chart_and_image(api, fixture_path):
    parser_module, _, _ = api
    _add_worksheet_chart_and_image(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def _rename_chart_part(path: Path, *, malformed: bool = False) -> None:
    _add_worksheet_chart(path)
    _rename_zip_member(
        path,
        "xl/charts/chart1.xml",
        "xl/charts/chart1.bin",
        (lambda data: data[:-1]) if malformed else (lambda data: data),
    )
    _rewrite_zip_member(
        path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(b"/xl/charts/chart1.xml", b"/xl/charts/chart1.bin"),
    )
    _rewrite_zip_member(
        path,
        "[Content_Types].xml",
        lambda data: data.replace(b"/xl/charts/chart1.xml", b"/xl/charts/chart1.bin"),
    )


@pytest.mark.parametrize(
    "relationship_namespace",
    [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ],
    ids=("transitional", "strict"),
)
def test_preflight_classifies_malformed_chart_by_relationship_not_suffix(
    api, fixture_path, monkeypatch, relationship_namespace
):
    parser_module, _, _ = api
    _rename_chart_part(fixture_path, malformed=True)
    if "purl.oclc.org" in relationship_namespace:
        _rewrite_zip_member(
            fixture_path,
            "xl/drawings/_rels/drawing1.xml.rels",
            lambda data: data.replace(
                b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                f"{relationship_namespace}/chart".encode(),
            ),
        )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Related XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_applies_xml_limit_to_chart_with_binary_suffix(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rename_chart_part(fixture_path)
    limit = 16 * 1024
    _rewrite_zip_member(
        fixture_path,
        "xl/charts/chart1.bin",
        lambda data: data.replace(b"</chartSpace>", b" " * limit + b"</chartSpace>"),
    )
    monkeypatch.setattr(
        parser_module, "MAX_PREFLIGHT_RELATED_XML_BYTES", limit
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"related XML size limit.*16384"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_well_formed_chart_with_binary_suffix(api, fixture_path):
    parser_module, _, _ = api
    _rename_chart_part(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_keeps_image_binary_despite_xml_suffix(api, fixture_path):
    parser_module, _, _ = api
    _add_worksheet_chart_and_image(fixture_path)
    _rename_zip_member(
        fixture_path, "xl/media/image1.png", "xl/media/image1.xml"
    )
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(b"/xl/media/image1.png", b"/xl/media/image1.xml"),
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/media/image1.xml" ContentType="image/png"/>'
            b"</Types>",
        ),
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_rejects_chart_consumer_disguised_as_image_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda data: data.replace(b"/relationships/chart\"", b"/relationships/image\""),
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
            b"image/png",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"content type conflicts with chart relationship semantics",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("semantic", "content_type"),
    [
        (
            "worksheet",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        ),
        (
            "chartsheet",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
        ),
        (
            "drawing",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
        ),
        (
            "table",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
        ),
        (
            "queryTable",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        ),
        (
            "pivotTable",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
        ),
        (
            "pivotCacheDefinition",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        ),
        (
            "pivotCacheRecords",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
        ),
        (
            "comments",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
        ),
        ("vmlDrawing", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
        ("chartStyle", "application/vnd.ms-office.chartstyle+xml"),
        ("chartColorStyle", "application/vnd.ms-office.chartcolorstyle+xml"),
    ],
)
@pytest.mark.parametrize(
    "relationship_namespace",
    [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ],
    ids=("transitional", "strict"),
)
def test_preflight_classifies_xml_relationship_semantics_for_both_namespaces(
    api, semantic, content_type, relationship_namespace
):
    parser_module, _, _ = api
    content_types = parser_module._ContentTypes(
        {}, {"xl/unusual.bin": content_type}, (("xl/unusual.bin", content_type),)
    )
    target = parser_module._RelationshipTarget(
        "xl/unusual.bin",
        "rId1",
        f"{relationship_namespace}/{semantic}",
        "xl/source.xml",
    )

    assert parser_module._validate_relationship_target_semantics(
        content_types, target
    ) is True


@pytest.mark.parametrize(
    "relationship_type",
    [
        "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
        "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
    ],
)
def test_preflight_rejects_microsoft_chart_style_content_type_conflict(
    api, relationship_type
):
    parser_module, _, _ = api
    content_types = parser_module._ContentTypes(
        {}, {"xl/style.bin": "image/png"}, (("xl/style.bin", "image/png"),)
    )
    target = parser_module._RelationshipTarget(
        "xl/style.bin", "rId1", relationship_type, "xl/charts/chart1.xml"
    )

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"content type conflicts with chart.*relationship semantics",
    ):
        parser_module._validate_relationship_target_semantics(
            content_types, target
        )


def test_preflight_rejects_relationship_content_type_conflict_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_worksheet_chart(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
            b"image/png",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"content type.*chart"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_content_type_mapping_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rename_chart_part(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b'<Override PartName="/xl/charts/chart1.bin" '
            b'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
            b"",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"content type mapping is missing"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_conflicting_content_type_overrides_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/styles.xml" ContentType="image/png"/>'
            b"</Types>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"conflicting content type override",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_content_types_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(fixture_path, "[Content_Types].xml", lambda data: data[:-1])
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Content types XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_wrong_content_types_root_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(b"<Types ", b"<NotTypes ", 1).replace(
            b"</Types>", b"</NotTypes>", 1
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Content types XML is invalid"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    "part_name", ["/xl/../escape.xml", "/xl//escape.xml", "/xl/./escape.xml"]
)
def test_preflight_rejects_unsafe_content_type_part_name_before_openpyxl(
    api, fixture_path, monkeypatch, part_name
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            f'<Override PartName="{part_name}" ContentType="application/xml"/>'.encode()
            + b"</Types>",
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"unsafe part name"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_content_type_record_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_CONTENT_TYPE_RECORDS", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"content type record limit.*1"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_missing_eager_style_target_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _remove_zip_member(fixture_path, "xl/styles.xml")
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"Relationship graph target is missing: xl/styles.xml",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_mismatched_office_document_target_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path,
        "_rels/.rels",
        lambda data: data.replace(b'Target="xl/workbook.xml"', b'Target="xl/styles.xml"'),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"officeDocument relationship target conflicts",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_unused_shared_string_records_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_zip_member(
        fixture_path,
        "xl/sharedStrings.bin",
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<si><t>one</t></si><si><t>two</t></si></sst>",
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/sharedStrings.bin" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            b"</Types>",
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_SHARED_STRING_RECORDS", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"shared string record limit.*1"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_shared_string_text_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_zip_member(
        fixture_path,
        "xl/sharedStrings.xml",
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<si><t>long text</t></si></sst>",
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/sharedStrings.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            b"</Types>",
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_SHARED_STRING_CHARACTERS", 4)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"shared string character limit.*4"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("run_xml", "description"),
    (
        (b"<r/>", "rich text"),
        (b'<rPh sb="0" eb="0"/>', "phonetic"),
    ),
)
def test_preflight_rejects_shared_string_run_records_before_openpyxl(
    api, fixture_path, monkeypatch, run_xml, description
):
    parser_module, _, _ = api
    record_limit = parser_module.MAX_PREFLIGHT_SHARED_STRING_OBJECT_RECORDS
    _add_shared_strings(
        fixture_path,
        b"<si>" + run_xml * (record_limit + 1) + b"</si>",
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=rf"shared string object limit.*{record_limit}",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_exact_shared_string_object_limit_with_rich_text(
    api, fixture_path
):
    parser_module, _, _ = api
    record_limit = parser_module.MAX_PREFLIGHT_SHARED_STRING_OBJECT_RECORDS
    _add_shared_strings(
        fixture_path,
        b"<si><r><t>ordinary text</t></r>"
        + b"<r/>" * (record_limit - 2)
        + b"</si>",
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_ignores_unrelated_rich_text_elements_in_shared_strings(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_shared_strings(
        fixture_path,
        b'<si><t>ordinary text</t><extLst><ext uri="urn:test">'
        b'<w:r xmlns:w="urn:unrelated"/></ext></extLst></si>',
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_SHARED_STRING_RECORDS", 1)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_accepts_bounded_shared_strings_with_unusual_name(
    api, fixture_path
):
    parser_module, _, _ = api
    _add_zip_member(
        fixture_path,
        "xl/sharedStrings.bin",
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<si><t>ordinary text</t></si></sst>",
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/xl/sharedStrings.bin" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            b"</Types>",
        ),
    )

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_rejects_style_record_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_STYLE_RECORDS", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"style record limit.*1"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_style_collection_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_STYLE_COLLECTION_RECORDS", 1)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"style collection record limit.*1",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_styles_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path, "xl/styles.xml", lambda data: data.rsplit(b">", 1)[0]
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Styles XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_custom_properties_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_zip_member(
        fixture_path,
        "docProps/custom.xml",
        b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">',
    )
    _rewrite_zip_member(
        fixture_path,
        "[Content_Types].xml",
        lambda data: data.replace(
            b"</Types>",
            b'<Override PartName="/docProps/custom.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
            b"</Types>",
        ),
    )
    _rewrite_zip_member(
        fixture_path,
        "_rels/.rels",
        lambda data: data.replace(
            b"</Relationships>",
            b'<Relationship Id="rIdCustom" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" '
            b'Target="docProps/custom.xml"/></Relationships>',
        ),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Package XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    "member_name",
    [
        "xl/theme/theme1.xml",
        "docProps/core.xml",
        "docProps/app.xml",
    ],
    ids=("theme", "core-properties", "extended-properties"),
)
def test_preflight_rejects_malformed_package_xml_before_openpyxl(
    api, fixture_path, monkeypatch, member_name
):
    parser_module, _, _ = api
    _rewrite_zip_member(
        fixture_path, member_name, lambda data: data.rsplit(b">", 1)[0]
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Package XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize(
    ("reference", "message"),
    [
        ("A20001", r"preflight row limit.*20000"),
        ("SS1", r"preflight column limit.*512"),
        ("A1:B2", r"invalid cell reference"),
    ],
    ids=("row", "column", "range"),
)
def test_preflight_rejects_invalid_comment_reference_before_openpyxl(
    api, fixture_path, monkeypatch, reference, message
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/comments/comment1.xml",
        lambda data: data.replace(b'ref="A1"', f'ref="{reference}"'.encode(), 1),
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_out_of_bounds_comment_on_irrelevant_sheet(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook.create_sheet("Irrelevant")["A20001"].comment = Comment("note", "author")
    workbook.save(fixture_path)
    workbook.close()
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"preflight row limit.*20000"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_rejects_malformed_comment_xml_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/comments/comment1.xml",
        lambda data: data[:-1],
    )
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"Comment XML is malformed"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_charges_duplicate_comments_to_sheet_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    _rewrite_zip_member(
        fixture_path,
        "xl/comments/comment1.xml",
        lambda data: data.replace(
            b"</commentList>",
            b'<comment ref="A1" authorId="0"><text><t>duplicate</t></text></comment>'
            b"</commentList>",
        ),
    )
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_WORKSHEET_CELLS", 84)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"cell materialization limit.*84",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_charges_comments_to_workbook_budget_before_openpyxl(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    _add_comment(fixture_path)
    monkeypatch.setattr(parser_module, "MAX_PREFLIGHT_TOTAL_CELLS", 123)
    monkeypatch.setattr(parser_module, "load_workbook", _fail_if_openpyxl_loads)

    with pytest.raises(
        parser_module.WorkbookStructureError,
        match=r"workbook cell materialization limit.*123",
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_preflight_accepts_normal_comment(api, fixture_path):
    parser_module, _, _ = api
    _add_comment(fixture_path)

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_preflight_accepts_external_hyperlink_relationship(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET]["A1"].hyperlink = "https://example.test/report"
    workbook.save(fixture_path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    assert len(rows) == 5


def test_parser_caps_total_entries_across_workbook(
    api, fixture_path, monkeypatch
):
    parser_module, _, _ = api
    monkeypatch.setattr(parser_module, "MAX_PARSED_ENTRIES", 4, raising=False)

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"entry limit.*4"
    ):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_rejects_regular_sheet_month_mismatch(api, tmp_path):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(tmp_path / "august.xlsx")

    with pytest.raises(parser_module.WorkbookStructureError, match="sheet month"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-09")


def test_parser_accepts_previous_calendar_month_as_august_billing(api, tmp_path):
    parser_module, _, _ = api
    path = build_transport_fixture(tmp_path / "previous-month.xlsx")
    workbook = load_workbook(path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 1, date(2026, 7, 31))
    workbook.save(path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    billed = next(
        row
        for row in rows
        if row.source_sheet == SUBCONTRACT_SHEETS[0] and row.source_row == 2
    )
    assert billed.report_month == "2026-08"
    assert billed.day == 31
    assert billed.source_date == date(2026, 7, 31)


def test_parser_accepts_previous_calendar_month_across_year_boundary(api, tmp_path):
    parser_module, _, _ = api
    path = build_transport_fixture(tmp_path / "january.xlsx", month=1)
    workbook = load_workbook(path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 1, date(2025, 12, 31))
    workbook.save(path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-01")

    billed = next(
        row
        for row in rows
        if row.source_sheet == SUBCONTRACT_SHEETS[0] and row.source_row == 2
    )
    assert billed.report_month == "2026-01"
    assert billed.source_date == date(2025, 12, 31)


@pytest.mark.parametrize("outside_date", [date(2026, 6, 30), date(2026, 9, 1)])
def test_parser_rejects_subcontract_dates_outside_billing_window(
    api, tmp_path, outside_date
):
    parser_module, _, _ = api
    path = build_transport_fixture(tmp_path / "outside-window.xlsx")
    workbook = load_workbook(path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 1, outside_date)
    workbook.save(path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="billing window"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_preserves_full_regular_source_date(api, fixture_path):
    parser_module, _, _ = api

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    regular = next(
        row
        for row in rows
        if row.source_sheet == REGULAR_SHEET and row.source_row == 3
    )
    assert regular.source_date == date(2026, 8, 12)


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


def test_parser_rejects_subcontract_detail_missing_vehicle(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 8).value = None
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="vehicle"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


@pytest.mark.parametrize("candidate_column", [3, 4, 10])
def test_parser_rejects_partial_subcontract_candidate_across_full_footprint(
    api, fixture_path, candidate_column
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[SUBCONTRACT_SHEETS[0]]
    sheet.insert_rows(3)
    sheet.cell(3, candidate_column, "partial detail")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="date"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_checks_driver_column_when_its_header_is_blank(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[SUBCONTRACT_SHEETS[0]]
    sheet.cell(1, 10).value = None
    sheet.insert_rows(3)
    sheet.cell(3, 10, "partial driver")
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="date"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_skips_pre_numbered_blank_subcontract_template_rows(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "pre-numbered-templates.xlsx",
        include_subcontract_template_rows=True,
    )

    rows = parser_module.HwaseongWorkbookParser().parse(path, "2026-08")

    assert len(rows) == 5


def test_parser_rejects_pre_total_numbered_subcontract_row_with_uncached_formula(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "pre-total-uncached-formula.xlsx",
        include_subcontract_pre_total_uncached_formula_row=True,
    )

    with pytest.raises(parser_module.WorkbookStructureError, match=r"row 4.*date"):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


def test_parser_rejects_post_total_numbered_subcontract_row_with_uncached_formula(
    api, tmp_path
):
    parser_module, _, _ = api
    path = build_structural_clone_fixture(
        tmp_path / "post-total-uncached-formula.xlsx",
        include_subcontract_post_total_uncached_formula_row=True,
    )

    with pytest.raises(
        parser_module.WorkbookStructureError, match=r"row 5.*detail after.*total"
    ):
        parser_module.HwaseongWorkbookParser().parse(path, "2026-08")


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


@pytest.mark.parametrize(
    "invalid_count",
    ["10000.0001", "0.00001", "1e100000", "12345678901234567"],
)
def test_parser_rejects_abusive_trip_counts_before_cost_calculation(
    api, fixture_path, invalid_count
):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[REGULAR_SHEET]
    sheet.cell(3, 14, invalid_count)
    sheet.cell(3, 34, invalid_count)
    workbook.save(fixture_path)
    workbook.close()

    with pytest.raises(parser_module.WorkbookStructureError, match="trip count"):
        parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")


def test_parser_accepts_trip_count_upper_boundary(api, fixture_path):
    parser_module, _, _ = api
    workbook = load_workbook(fixture_path)
    sheet = workbook[REGULAR_SHEET]
    sheet.cell(3, 14, 10_000)
    sheet.cell(3, 34, 10_000)
    sheet.cell(3, 36, 1)
    sheet.cell(3, 37, 10_000)
    workbook.save(fixture_path)
    workbook.close()

    rows = parser_module.HwaseongWorkbookParser().parse(fixture_path, "2026-08")

    boundary = next(row for row in rows if row.source_row == 3)
    assert boundary.trip_count == Decimal("10000")


def test_repository_rejects_exponent_abuse_before_persistence(api, database):
    parser_module, repository_module, _ = api
    row = parser_module.ParsedTransportEntry(
        report_month="2026-08",
        destination_alias="Known Plant",
        source_sheet=REGULAR_SHEET,
        source_row=3,
        source_date=date(2026, 8, 1),
        day=1,
        transport_type="regular",
        trip_count=Decimal("1e100000"),
        unit_rate_won=1,
        cost_won=1,
    )

    with pytest.raises(parser_module.WorkbookStructureError, match="trip count"):
        repository_module.TransportEntryRepository(database).import_entries(
            report_month="2026-08",
            source_filename="synthetic.xlsx",
            file_sha256="e" * 64,
            rows=[row],
        )

    with database.connection() as connection:
        assert (
            connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0]
            == 0
        )
        assert (
            connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0]
            == 0
        )


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"report_month": "2026-07"}, "report_month must match"),
        ({"day": 2}, "day must match source_date"),
        ({"source_date": None}, "source_date is required"),
        ({"transport_type": "charter"}, "transport_type"),
        (
            {"source_date": date(2026, 7, 31), "day": 31},
            "regular source_date",
        ),
        (
            {
                "source_date": date(2026, 6, 30),
                "day": 30,
                "transport_type": "nonregular",
                "unit_rate_won": None,
            },
            "nonregular source_date",
        ),
        (
            {
                "source_date": date(2026, 9, 1),
                "transport_type": "nonregular",
                "unit_rate_won": None,
            },
            "nonregular source_date",
        ),
    ],
    ids=(
        "wrong-row-month",
        "wrong-day",
        "missing-source-date",
        "invalid-type",
        "regular-prior-month",
        "nonregular-too-old",
        "nonregular-future",
    ),
)
def test_repository_rejects_inconsistent_entries_before_commit(
    api, database, overrides, message
):
    parser_module, repository_module, _ = api
    row = _parsed_entry(parser_module, **overrides)

    with pytest.raises(parser_module.WorkbookStructureError, match=message):
        repository_module.TransportEntryRepository(database).import_entries(
            report_month="2026-08",
            source_filename="synthetic.xlsx",
            file_sha256="d" * 64,
            rows=[row],
        )

    with database.connection() as connection:
        assert connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0] == 0
        assert connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0] == 0


def test_repository_accepts_nonregular_previous_month_across_year_boundary(
    api, database
):
    parser_module, repository_module, _ = api
    row = _parsed_entry(
        parser_module,
        report_month="2026-01",
        source_date=date(2025, 12, 31),
        day=31,
        transport_type="nonregular",
        unit_rate_won=None,
    )

    result = repository_module.TransportEntryRepository(database).import_entries(
        report_month="2026-01",
        source_filename="synthetic.xlsx",
        file_sha256="c" * 64,
        rows=[row],
    )

    assert result.status == "imported"
    assert result.inserted_count == 1


def test_duplicate_import_restores_newline_alias_blockers_from_entries(api, database):
    parser_module, repository_module, _ = api
    rows = [
        parser_module.ParsedTransportEntry(
            report_month="2026-08",
            destination_alias=alias,
            source_sheet=REGULAR_SHEET,
            source_row=source_row,
            source_date=date(2026, 8, 1),
            day=1,
            transport_type="regular",
            trip_count=Decimal("1"),
            unit_rate_won=1,
            cost_won=1,
        )
        for source_row, alias in enumerate(
            ("Unknown\nPlant", "Unknown\nPlant", "Other Plant"), start=3
        )
    ]
    repository = repository_module.TransportEntryRepository(database)

    first = repository.import_entries(
        report_month="2026-08",
        source_filename="synthetic.xlsx",
        file_sha256="f" * 64,
        rows=rows,
    )
    duplicate = repository.import_entries(
        report_month="2026-08",
        source_filename="synthetic.xlsx",
        file_sha256="f" * 64,
        rows=rows,
    )

    expected_errors = (
        "Unknown destination alias: Unknown\nPlant",
        "Unknown destination alias: Other Plant",
    )
    assert first.status == "imported_with_errors"
    assert first.blocking_errors == expected_errors
    assert duplicate.status == "duplicate"
    assert duplicate.inserted_count == 0
    assert duplicate.blocking_errors == expected_errors
    with database.connection() as connection:
        assert (
            connection.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0]
            == 1
        )
        assert (
            connection.execute("SELECT COUNT(*) FROM transport_entries").fetchone()[0]
            == 3
        )


def test_import_is_idempotent_and_stores_file_hash(api, database, fixture_path):
    _, _, service_module = api
    workbook = load_workbook(fixture_path)
    workbook[REGULAR_SHEET].cell(5, 2, "Known Plant")
    workbook.save(fixture_path)
    workbook.close()
    service = service_module.TransportImportService(database)

    first = service.import_transport(fixture_path, "2026-08")
    second = service.import_transport(fixture_path, "2026-08")

    assert first.status == "imported"
    assert first.inserted_count == 5
    assert second.status == "duplicate"
    assert second.inserted_count == 0
    assert second.blocking_errors == ()
    with database.connection() as connection:
        batches = connection.execute("SELECT * FROM import_batches").fetchall()
        entries = connection.execute("SELECT * FROM transport_entries").fetchall()
    assert len(batches) == 1
    assert len(entries) == 5
    assert batches[0]["file_sha256"] == hashlib.sha256(
        fixture_path.read_bytes()
    ).hexdigest()


def test_import_persists_complete_source_dates(api, database, tmp_path):
    _, _, service_module = api
    path = build_transport_fixture(tmp_path / "source-dates.xlsx")
    workbook = load_workbook(path)
    workbook[SUBCONTRACT_SHEETS[0]].cell(2, 1, date(2026, 7, 31))
    workbook.save(path)
    workbook.close()

    service_module.TransportImportService(database).import_transport(path, "2026-08")

    with database.connection() as connection:
        source_dates = {
            row["source_date"]
            for row in connection.execute(
                "SELECT source_date FROM transport_entries"
            )
        }
    assert "2026-07-31" in source_dates
    assert "2026-08-12" in source_dates


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
    service = service_module.TransportImportService(database)

    result = service.import_transport(fixture_path, "2026-08")
    destination = MasterRepository(database).list_destinations()[0]
    MasterRepository(database).add_alias(destination.id, "Unknown Plant", "transport")
    duplicate = service.import_transport(fixture_path, "2026-08")

    assert result.status == "imported_with_errors"
    assert result.blocking_errors == ("Unknown destination alias: Unknown Plant",)
    assert duplicate.status == "duplicate"
    assert duplicate.inserted_count == 0
    assert duplicate.blocking_errors == result.blocking_errors
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
