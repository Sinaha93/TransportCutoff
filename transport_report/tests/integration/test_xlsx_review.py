"""Fictional integration coverage for the review-workbook exporter."""
from __future__ import annotations

from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook


ERROR_TOKENS = {
    "#REF!",
    "#DIV/0!",
    "#VALUE!",
    "#NAME?",
    "#N/A",
    "#NUM!",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
}


def _report_bundle():
    from app.domain.calculations import calculate_destination, calculate_group, calculate_total
    from app.domain.models import Destination, DestinationAlias, GroupMember, ReportGroup
    from app.domain.validation import ValidationIssue, ValidationResult
    from app.reporting.charts import ChartReport
    from app.reporting.pptx_report import PlanReport, PptReport, ReportRow, SalesReport
    from app.reporting.xlsx_review import (
        ImportBatchEvidence,
        OperationEvidence,
        ReviewWorkbookReport,
    )

    destinations = (
        Destination(1, "가상동부", 1, True, True, "가상품 A", True, True, True),
        Destination(2, "가상서부", 2, True, True, None, True, True, False),
    )
    aliases = (
        DestinationAlias(1, 1, "동부 원천", "transport"),
        DestinationAlias(2, 2, "서부 원천", "transport"),
    )
    group = ReportGroup(10, "가상권역", 1, True)
    members = (
        GroupMember(10, 1, "가상동부", 1, True, True),
        GroupMember(10, 2, "가상서부", 2, True, False),
    )
    calculations = {
        1: calculate_destination(Decimal("100"), 10_000, Decimal("120"), 13_200, destination_id=1),
        2: calculate_destination(Decimal("50"), 5_000, Decimal("0"), 0, destination_id=2),
    }
    group_calculation = calculate_group(calculations, members, group_id=10)
    total_calculation = calculate_total(calculations, destinations)
    months = tuple(f"2025-{month:02d}" for month in range(1, 13)) + tuple(
        f"2026-{month:02d}" for month in range(1, 9)
    )

    def row(key, label, calculation):
        quantities = {month: Decimal("90") for month in months}
        costs = {month: 9_000 for month in months}
        quantities["2026-08"] = calculation.actual_quantity
        costs["2026-08"] = calculation.actual_cost_won
        return ReportRow(key, label, calculation, quantities, costs)

    first = row("1", "가상동부", calculations[1])
    second = row("2", "가상서부", calculations[2])
    group_row = row("group:10", "가상권역", group_calculation)
    total_row = row("total", "합계", total_calculation)
    nonregular = row(
        "nonregular",
        "비정규 운반비",
        calculate_destination(Decimal("0"), 0, Decimal("0"), 0),
    )
    sales_history = {month: 1_000_000 for month in months}
    sales_history["2026-08"] = 1_200_000
    sales = SalesReport(1_100_000, 1_200_000, sales_history)
    next_rows = (
        replace(first, calculation=calculate_destination(Decimal("130"), 14_300, None, None, destination_id=1)),
        replace(second, calculation=calculate_destination(Decimal("60"), 6_600, None, None, destination_id=2)),
    )
    next_total = calculate_total({1: next_rows[0].calculation, 2: next_rows[1].calculation}, destinations)
    chart = ChartReport(
        "2026-08",
        {month: Decimal("150") for month in months},
        {month: total_row.quantity_by_month[month] for month in months},
        {month: 15_000 for month in months},
        {month: total_row.cost_won_by_month[month] for month in months},
    )
    report = PptReport(
        "2026-08",
        date(2026, 9, 2),
        (first, second, group_row),
        nonregular,
        total_row,
        sales,
        PlanReport(
            "2026-09",
            next_rows,
            replace(nonregular, calculation=calculate_destination(Decimal("0"), 0, None, None)),
            replace(total_row, calculation=next_total),
            SalesReport(1_300_000, None, sales_history),
        ),
        chart,
    )
    batch = ImportBatchEvidence(
        batch_id=42,
        provenance_id="transport:2026-08:42",
        report_month="2026-08",
        source_type="transport",
        source_filename=r"D:\\fictional\\화성 가상자료.xlsx",
        file_sha256="ab" * 32,
        imported_at=datetime(2026, 9, 1, 8, 30, 15),
    )
    operations = (
        OperationEvidence(
            input_kind="imported",
            batch_id=42,
            provenance_id="transport:2026-08:42",
            report_month="2026-08",
            raw_destination="동부 원천",
            destination_id=1,
            normalized_destination="가상동부",
            quantity_ea=Decimal("120"),
            cost_won=13_200,
            source_locator=r"D:\fictional\화성운반비내역(8월)!12",
        ),
        OperationEvidence(
            input_kind="manual",
            batch_id=None,
            provenance_id="monthly-actual:2026-08:2",
            report_month="2026-08",
            raw_destination=None,
            destination_id=2,
            normalized_destination="가상서부",
            quantity_ea=Decimal("0"),
            cost_won=None,
            source_locator="당월 실적 입력",
        ),
        OperationEvidence(
            input_kind="imported",
            batch_id=42,
            provenance_id="transport:2026-08:42",
            report_month="2026-08",
            raw_destination="서부 원천",
            destination_id=2,
            normalized_destination="가상서부",
            quantity_ea=None,
            cost_won=0,
            source_locator="화성운반비내역(8월)!13",
        ),
    )
    validation = ValidationResult(
        (
            ValidationIssue(
                code="CHECK_NOTE",
                severity="warning",
                message="가상 검토 메모입니다.",
                destination_id=2,
                source_locator="당월 실적 입력",
            ),
        )
    )
    return ReviewWorkbookReport(
        report=report,
        validation=validation,
        destinations=destinations,
        aliases=aliases,
        groups=(group,),
        group_members=members,
        import_batches=(batch,),
        operations=operations,
    )


def _find_row(sheet, label: str) -> int:
    for row in sheet.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise AssertionError(f"{label!r} row not found in {sheet.title}")


def _as_date(value):
    return value.date() if isinstance(value, datetime) else value


def test_export_review_workbook_reconciles_typed_values_and_review_evidence(tmp_path):
    from app.reporting.xlsx_review import export_review_workbook

    bundle = _report_bundle()
    output = tmp_path / "review.xlsx"
    export_review_workbook(bundle, output)

    workbook = load_workbook(output, data_only=False, keep_links=True)
    assert workbook.sheetnames == ["월간 종합", "운행실적", "검증 결과", "마스터 기준"]
    assert workbook.vba_archive is None
    assert not workbook._external_links

    summary = workbook["월간 종합"]
    assert summary.sheet_view.showGridLines is False
    assert summary.freeze_panes == "C9"
    assert summary.auto_filter.ref
    assert _as_date(summary["B3"].value) == date(2026, 8, 1)
    assert summary["B3"].is_date
    total_row = _find_row(summary, "합계")
    total = bundle.report.total.calculation
    assert summary.cell(total_row, 3).value == float(total.planned_quantity)
    assert summary.cell(total_row, 4).value == float(total.actual_quantity)
    assert summary.cell(total_row, 7).value == total.planned_cost_won
    assert summary.cell(total_row, 8).value == total.actual_cost_won
    assert summary.cell(total_row, 6).data_type == "n"
    assert summary.cell(total_row, 6).number_format == "0.0%"
    zero_row = _find_row(summary, "가상서부")
    assert summary.cell(zero_row, 4).value == 0
    assert summary.cell(zero_row, 12).value is None
    assert summary.cell(zero_row, 14).value is None
    assert summary.cell(zero_row, 4).fill.fgColor.rgb == summary.cell(zero_row, 3).fill.fgColor.rgb
    assert summary.cell(zero_row, 8).fill.fgColor.rgb == summary.cell(9, 8).fill.fgColor.rgb

    sales_row = _find_row(summary, "매출액")
    assert summary.cell(sales_row, 3).value == 1_100_000
    assert summary.cell(sales_row, 4).value == 1_200_000
    ratio_row = _find_row(summary, "운반비/매출")
    assert isinstance(summary.cell(ratio_row, 4).value, (int, float))
    assert summary.cell(ratio_row, 4).number_format == "0.0%"
    averages_row = _find_row(summary, "수량(EA)")
    assert isinstance(summary.cell(averages_row, 2).value, (int, float))
    assert summary.cell(averages_row, 3).value == "완전"

    evidence = workbook["운행실적"]
    assert evidence.freeze_panes == "A10"
    assert evidence.auto_filter.ref
    assert evidence["B4"].value == "화성 가상자료.xlsx"
    assert evidence["B5"].value == "ab" * 32
    assert evidence["B6"].value == datetime(2026, 9, 1, 8, 30, 15)
    assert evidence["B6"].is_date
    assert evidence["B7"].value == 42
    assert _as_date(evidence["D10"].value) == date(2026, 8, 1)
    assert evidence["D10"].is_date
    assert evidence["A10"].value == "가져오기"
    assert evidence["I10"].value == "화성운반비내역(8월)!12"
    assert "D:\\fictional" not in evidence["I10"].value
    assert evidence["A11"].value == "수기 입력"
    assert evidence["G11"].value == 0
    assert evidence["H11"].value is None
    assert evidence["H12"].value == 0
    assert evidence["A10"].fill.fgColor.rgb != evidence["A11"].fill.fgColor.rgb

    checks = workbook["검증 결과"]
    assert checks["B3"].value == "가능"
    assert checks["B4"].value == 0
    assert checks.freeze_panes == "A8"
    assert checks.auto_filter.ref
    assert checks["B8"].value == "경고"
    assert checks["B8"].fill.fgColor.rgb != "00000000"

    masters = workbook["마스터 기준"]
    assert masters.protection.sheet is True
    assert masters.freeze_panes == "A6"
    assert masters.auto_filter.ref
    assert masters["B6"].value == "가상동부"
    assert "동부 원천 (transport)" in masters["C6"].value
    group_header = _find_row(masters, "그룹명")
    assert masters.cell(group_header + 1, 1).value == "가상권역"

    for sheet in workbook.worksheets:
        assert any(cell.style_id for row in sheet.iter_rows() for cell in row)
        assert all(dimension.width and dimension.width < 60 for dimension in sheet.column_dimensions.values())
        for row in sheet.iter_rows():
            for cell in row:
                assert not (isinstance(cell.value, str) and cell.value in ERROR_TOKENS)
                assert cell.data_type != "e"
                assert not (isinstance(cell.value, str) and cell.value.startswith("="))


def test_export_review_workbook_shows_neutral_valid_state(tmp_path):
    from app.domain.validation import ValidationResult
    from app.reporting.xlsx_review import export_review_workbook

    bundle = replace(_report_bundle(), validation=ValidationResult(()))
    output = tmp_path / "valid.xlsx"
    export_review_workbook(bundle, output)

    checks = load_workbook(output)["검증 결과"]
    assert checks["B3"].value == "가능"
    assert checks["B4"].value == 0
    assert checks["A8"].value == "VALID"
    assert checks["B8"].value == "정상"
    assert checks["A8"].fill.fgColor.rgb in {"00000000", "00FFFFFF"}


def test_export_review_workbook_shows_blocking_errors_distinctly(tmp_path):
    from app.domain.validation import ValidationIssue, ValidationResult
    from app.reporting.xlsx_review import export_review_workbook

    bundle = replace(
        _report_bundle(),
        validation=ValidationResult(
            (
                ValidationIssue(
                    code="MISSING_SALES",
                    severity="error",
                    message="가상 매출 입력이 없습니다.",
                    source_locator="월 매출 입력",
                ),
            )
        ),
    )
    output = tmp_path / "blocked.xlsx"
    export_review_workbook(bundle, output)

    checks = load_workbook(output)["검증 결과"]
    assert checks["B3"].value == "불가"
    assert checks["B4"].value == 1
    assert checks["B8"].value == "오류"
    assert checks["B8"].fill.fgColor.rgb != "00000000"


@pytest.mark.parametrize("case", ["duplicate", "missing"])
def test_export_review_workbook_rejects_duplicate_or_missing_report_identity(tmp_path, case):
    from app.domain.calculations import calculate_destination
    from app.reporting.xlsx_review import export_review_workbook

    bundle = _report_bundle()
    first = bundle.report.rows[0]
    if case == "duplicate":
        rows = (first, first)
    else:
        rows = (
            replace(
                first,
                calculation=calculate_destination(
                    Decimal("100"), 10_000, Decimal("120"), 13_200
                ),
            ),
        )
    invalid = replace(bundle, report=replace(bundle.report, rows=rows))
    output = tmp_path / f"{case}.xlsx"

    with pytest.raises(ValueError, match="identity"):
        export_review_workbook(invalid, output)

    assert not output.exists()


@pytest.mark.parametrize(
    ("case", "message"),
    [
        ("alias", "alias identity"),
        ("evidence", "evidence does not reconcile"),
        ("duplicate-evidence", "duplicate identity"),
        ("batch-month", "batch report month"),
        ("group", "group calculation"),
        ("total", "total calculation"),
    ],
)
def test_export_review_workbook_rejects_inconsistent_audit_provenance(
    tmp_path, case, message
):
    from app.domain.calculations import calculate_group, calculate_total
    from app.reporting.xlsx_review import export_review_workbook

    bundle = _report_bundle()
    if case == "alias":
        first = replace(bundle.operations[0], raw_destination="미등록 원천")
        invalid = replace(bundle, operations=(first, *bundle.operations[1:]))
    elif case == "evidence":
        first = replace(bundle.operations[0], quantity_ea=Decimal("119"))
        invalid = replace(bundle, operations=(first, *bundle.operations[1:]))
    elif case == "duplicate-evidence":
        invalid = replace(bundle, operations=(*bundle.operations, bundle.operations[0]))
    elif case == "batch-month":
        invalid = replace(
            bundle,
            import_batches=(replace(bundle.import_batches[0], report_month="2026-07"),),
        )
    elif case == "group":
        calculations = {
            row.calculation.provenance.destination_id: row.calculation
            for row in bundle.report.rows[:2]
            if row is not None
        }
        wrong = calculate_group(calculations, bundle.group_members[:1], group_id=10)
        rows = (*bundle.report.rows[:2], replace(bundle.report.rows[2], calculation=wrong))
        invalid = replace(bundle, report=replace(bundle.report, rows=rows))
    else:
        direct_rows = tuple(row for row in bundle.report.rows if row is not None)[:2]
        calculations = {
            row.calculation.provenance.destination_id: row.calculation
            for row in direct_rows
        }
        wrong = calculate_total(calculations, bundle.destinations[:1])
        invalid = replace(
            bundle,
            report=replace(
                bundle.report,
                total=replace(bundle.report.total, calculation=wrong),
            ),
        )
    output = tmp_path / f"{case}.xlsx"

    with pytest.raises(ValueError, match=message):
        export_review_workbook(invalid, output)

    assert not output.exists()


def test_export_review_workbook_is_atomic_when_final_replace_fails(tmp_path, monkeypatch):
    from app.reporting import xlsx_review

    output = tmp_path / "existing.xlsx"
    original = b"existing review workbook"
    output.write_bytes(original)

    def fail_replace(source: Path, destination: Path) -> None:
        raise OSError("synthetic replace failure")

    monkeypatch.setattr(xlsx_review.os, "replace", fail_replace)
    with pytest.raises(OSError, match="synthetic replace failure"):
        xlsx_review.export_review_workbook(_report_bundle(), output)

    assert output.read_bytes() == original
    assert list(tmp_path.glob(".*.tmp.xlsx")) == []
