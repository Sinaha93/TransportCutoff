from decimal import Decimal

import pytest


def _complete_prior_year(report_month: str = "2026-08"):
    from app.domain.calculations import AverageResult

    year = int(report_month[:4]) - 1
    months = tuple(f"{year:04d}-{month:02d}" for month in range(1, 13))
    return AverageResult(
        expected_months=months,
        missing_months=(),
        value=Decimal("1"),
        complete=True,
    )


def _context(**changes):
    from app.domain.validation import SalesInput, ValidationContext

    values = {
        "report_month": "2026-08",
        "sales": SalesInput(
            amount_won=1_000_000,
            confirmed_at="2026-08-31",
            source_locator="월 매출 입력",
        ),
        "prior_year_history": _complete_prior_year(),
    }
    values.update(changes)
    return ValidationContext(**values)


def test_unknown_destination_alias_reports_each_source_row_as_blocking_error():
    from app.domain.validation import UnresolvedDestinationAlias, validate_report

    context = _context(
        unresolved_aliases=(
            UnresolvedDestinationAlias("미등록 납품처", "정규운송!37행"),
            UnresolvedDestinationAlias("미등록 납품처", "용차!12행"),
        )
    )

    result = validate_report(context)

    assert [issue.code for issue in result.issues] == [
        "MISSING_DESTINATION_ALIAS",
        "MISSING_DESTINATION_ALIAS",
    ]
    assert [issue.source_locator for issue in result.issues] == [
        "용차!12행",
        "정규운송!37행",
    ]
    assert all(issue.severity == "error" for issue in result.issues)
    assert "미등록 납품처" in result.issues[0].message
    assert "별칭" in result.issues[0].message
    assert result.can_generate is False


def test_missing_required_vehicle_rate_is_blocking_and_names_the_vehicle():
    from app.domain.validation import DestinationValidationInput, validate_report

    context = _context(
        destinations=(
            DestinationValidationInput(
                destination_id=7,
                name="울산공장",
                display_order=2,
                required_for_report=True,
                actual_quantity=Decimal("10"),
                next_month_plan_quantity=Decimal("11"),
                required_vehicle_types=("5톤", "11톤"),
                rated_vehicle_types=("5톤",),
            ),
        )
    )

    result = validate_report(context)

    issue = next(
        issue for issue in result.issues if issue.code == "MISSING_VEHICLE_RATE"
    )
    assert issue.severity == "error"
    assert issue.destination_id == 7
    assert issue.source_locator == "차량 요율 기준정보"
    assert "울산공장" in issue.message
    assert "11톤" in issue.message
    assert "등록" in issue.message


def test_required_destination_missing_actual_quantity_is_blocking():
    from app.domain.validation import DestinationValidationInput, validate_report

    context = _context(
        destinations=(
            DestinationValidationInput(
                destination_id=3,
                name="영천공장",
                display_order=1,
                required_for_report=True,
                actual_quantity=None,
                next_month_plan_quantity=Decimal("1"),
            ),
        )
    )

    result = validate_report(context)

    issue = next(
        issue for issue in result.issues if issue.code == "MISSING_ACTUAL_QUANTITY"
    )
    assert issue.severity == "error"
    assert issue.destination_id == 3
    assert issue.source_locator == "당월 실적 입력"
    assert "영천공장" in issue.message
    assert "실적 수량" in issue.message


def test_explicit_zero_actual_quantity_is_present():
    from app.domain.validation import DestinationValidationInput, validate_report

    context = _context(
        destinations=(
            DestinationValidationInput(
                destination_id=3,
                name="영천공장",
                display_order=1,
                required_for_report=True,
                actual_quantity=Decimal("0"),
                next_month_plan_quantity=Decimal("0"),
            ),
        )
    )

    result = validate_report(context)

    assert "MISSING_ACTUAL_QUANTITY" not in {issue.code for issue in result.issues}
    assert "MISSING_NEXT_MONTH_PLAN" not in {issue.code for issue in result.issues}


def test_missing_monthly_sales_amount_is_blocking():
    from app.domain.validation import SalesInput, validate_report

    result = validate_report(
        _context(
            sales=SalesInput(
                amount_won=None,
                confirmed_at="2026-08-31",
                source_locator="월 매출 입력",
            )
        )
    )

    issue = next(issue for issue in result.issues if issue.code == "MISSING_SALES")
    assert issue.severity == "error"
    assert issue.source_locator == "월 매출 입력"
    assert "매출액" in issue.message
    assert "확정일" in issue.message
    assert "입력" in issue.message


def test_sales_without_confirmation_date_is_still_missing_sales():
    from app.domain.validation import SalesInput, validate_report

    result = validate_report(
        _context(
            sales=SalesInput(
                amount_won=0,
                confirmed_at=None,
                source_locator="월 매출 입력",
            )
        )
    )

    issue = next(issue for issue in result.issues if issue.code == "MISSING_SALES")
    assert issue.severity == "error"
    assert "확정일" in issue.message
    assert result.can_generate is False


def test_conflicting_current_import_batches_are_blocking_duplicates():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _context(
            current_import_batches=(
                ImportBatchInput(21, "2026-08", "transport", "sha-a", "배치 #21"),
                ImportBatchInput(22, "2026-08", "transport", "sha-b", "배치 #22"),
            )
        )
    )

    issue = next(issue for issue in result.issues if issue.code == "DUPLICATE_IMPORT")
    assert issue.severity == "error"
    assert issue.source_locator == "배치 #21, 배치 #22"
    assert "서로 다른" in issue.message
    assert "하나만" in issue.message
    assert "transport" not in issue.message


def test_same_file_import_retry_is_idempotent_not_a_duplicate_conflict():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _context(
            current_import_batches=(
                ImportBatchInput(21, "2026-08", "transport", "same-sha", "배치 #21"),
                ImportBatchInput(21, "2026-08", "transport", "same-sha", "재시도"),
            )
        )
    )

    assert "DUPLICATE_IMPORT" not in {issue.code for issue in result.issues}


def test_duplicate_batch_locator_order_is_stable_when_batch_ids_tie():
    from app.domain.validation import ImportBatchInput, validate_report

    batches = (
        ImportBatchInput(21, "2026-08", "transport", "sha-z", "배치 Z"),
        ImportBatchInput(21, "2026-08", "transport", "sha-a", "배치 A"),
    )

    forward = validate_report(_context(current_import_batches=batches))
    reversed_result = validate_report(
        _context(current_import_batches=tuple(reversed(batches)))
    )

    forward_issue = next(
        issue for issue in forward.issues if issue.code == "DUPLICATE_IMPORT"
    )
    reversed_issue = next(
        issue for issue in reversed_result.issues if issue.code == "DUPLICATE_IMPORT"
    )
    assert forward_issue == reversed_issue
    assert forward_issue.source_locator == "배치 A, 배치 Z"


def test_transport_subtotal_mismatch_is_blocking():
    from app.domain.validation import TransportSubtotalInput, validate_report

    result = validate_report(
        _context(
            transport_subtotals=(
                TransportSubtotalInput(
                    destination_id=7,
                    destination_name="울산공장",
                    detail_total_won=120_000,
                    reported_subtotal_won=119_000,
                    source_locator="정규운송!AK37",
                ),
            )
        )
    )

    issue = next(
        issue
        for issue in result.issues
        if issue.code == "TRANSPORT_SUBTOTAL_MISMATCH"
    )
    assert issue.severity == "error"
    assert issue.destination_id == 7
    assert issue.source_locator == "정규운송!AK37"
    assert "120,000" in issue.message
    assert "119,000" in issue.message
    assert "상세 행" in issue.message


def test_exact_destination_and_grand_total_reconciliation_is_required():
    from app.domain.validation import TotalReconciliationInput, validate_report

    result = validate_report(
        _context(
            total_reconciliations=(
                TotalReconciliationInput(
                    metric_label="실적 수량",
                    value_kind="quantity",
                    detail_total=Decimal("10.001"),
                    grand_total=Decimal("10.000"),
                    source_locator="보고서 총계",
                ),
            )
        )
    )

    issue = next(
        issue
        for issue in result.issues
        if issue.code == "TOTAL_RECONCILIATION_FAILED"
    )
    assert issue.severity == "error"
    assert issue.source_locator == "보고서 총계"
    assert "실적 수량" in issue.message
    assert "10.001" in issue.message
    assert "10.000" in issue.message
    assert "총계" in issue.message


def test_missing_next_month_plan_is_blocking():
    from app.domain.validation import DestinationValidationInput, validate_report

    result = validate_report(
        _context(
            destinations=(
                DestinationValidationInput(
                    destination_id=9,
                    name="당진공장",
                    display_order=4,
                    required_for_report=True,
                    actual_quantity=Decimal("0"),
                    next_month_plan_quantity=None,
                ),
            )
        )
    )

    issue = next(
        issue for issue in result.issues if issue.code == "MISSING_NEXT_MONTH_PLAN"
    )
    assert issue.severity == "error"
    assert issue.destination_id == 9
    assert issue.source_locator == "다음 달 계획 입력"
    assert "2026-09" in issue.message
    assert "계획" in issue.message
    assert result.can_generate is False


def test_incomplete_prior_calendar_year_history_lists_missing_months():
    from app.domain.calculations import AverageResult
    from app.domain.validation import validate_report

    expected = tuple(f"2025-{month:02d}" for month in range(1, 13))
    result = validate_report(
        _context(
            prior_year_history=AverageResult(
                expected_months=expected,
                missing_months=("2025-02", "2025-11"),
                value=None,
                complete=False,
            )
        )
    )

    issue = next(
        issue for issue in result.issues if issue.code == "MISSING_PRIOR_YEAR_HISTORY"
    )
    assert issue.severity == "error"
    assert issue.source_locator == "전년도 이력"
    assert "2025-02" in issue.message
    assert "2025-11" in issue.message
    assert "보완" in issue.message
    assert result.can_generate is False


def test_auxiliary_warning_does_not_block_generation():
    from app.domain.validation import ValidationIssue, ValidationResult

    result = ValidationResult(
        issues=(
            ValidationIssue(
                code="OPTIONAL_REVIEW_NOTE",
                severity="warning",
                message="참고용 검토 메모입니다.",
            ),
        )
    )

    assert result.can_generate is True


@pytest.mark.parametrize(
    "code",
    [
        "MISSING_DESTINATION_ALIAS",
        "MISSING_VEHICLE_RATE",
        "MISSING_ACTUAL_QUANTITY",
        "MISSING_SALES",
        "DUPLICATE_IMPORT",
        "TRANSPORT_SUBTOTAL_MISMATCH",
        "TOTAL_RECONCILIATION_FAILED",
        "MISSING_NEXT_MONTH_PLAN",
        "MISSING_PRIOR_YEAR_HISTORY",
        "REPORT_MONTH_MISMATCH",
    ],
)
def test_required_blocking_codes_cannot_be_downgraded_to_warning(code):
    from app.domain.validation import ValidationIssue

    with pytest.raises(ValueError, match="blocking"):
        ValidationIssue(code=code, severity="warning", message="검토하세요.")


@pytest.mark.parametrize(
    ("source_name", "source_locator"),
    [
        ("plan", "계획 입력"),
        ("actual", "실적 입력"),
        ("import", "배치 #21"),
        ("title", "표지 제목"),
        ("graph_last_month", "그래프 마지막 열"),
    ],
)
def test_every_report_month_source_must_match_configured_month(
    source_name, source_locator
):
    from app.domain.validation import ReportMonthInput, validate_report

    result = validate_report(
        _context(
            month_inputs=(
                ReportMonthInput(source_name, "2026-07", source_locator),
            )
        )
    )

    issue = next(
        issue for issue in result.issues if issue.code == "REPORT_MONTH_MISMATCH"
    )
    assert issue.severity == "error"
    assert issue.source_locator == source_locator
    assert "2026-07" in issue.message
    assert "2026-08" in issue.message
    assert "수정" in issue.message


def test_month_mismatch_does_not_hide_conflicting_current_batches():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _context(
            current_import_batches=(
                ImportBatchInput(31, "2026-07", "transport", "sha-a", "배치 #31"),
                ImportBatchInput(32, "2026-07", "transport", "sha-b", "배치 #32"),
            )
        )
    )

    codes = [issue.code for issue in result.issues]
    assert codes.count("REPORT_MONTH_MISMATCH") == 2
    assert codes.count("DUPLICATE_IMPORT") == 1


def test_issues_are_deduplicated_only_when_every_identifying_field_matches():
    from app.domain.validation import UnresolvedDestinationAlias, validate_report

    duplicate = UnresolvedDestinationAlias("미등록", "정규운송!10행")
    result = validate_report(
        _context(
            unresolved_aliases=(
                duplicate,
                duplicate,
                UnresolvedDestinationAlias("미등록", "정규운송!11행"),
            )
        )
    )

    alias_issues = [
        issue
        for issue in result.issues
        if issue.code == "MISSING_DESTINATION_ALIAS"
    ]
    assert [issue.source_locator for issue in alias_issues] == [
        "정규운송!10행",
        "정규운송!11행",
    ]


def test_issue_sorting_is_deterministic_by_severity_destination_code_and_source():
    from app.domain.validation import (
        DestinationValidationInput,
        TransportSubtotalInput,
        UnresolvedDestinationAlias,
        validate_report,
    )

    context = _context(
        destinations=(
            DestinationValidationInput(
                10,
                "후순위",
                5,
                True,
                None,
                Decimal("1"),
            ),
            DestinationValidationInput(
                20,
                "선순위",
                1,
                True,
                Decimal("1"),
                None,
                ("11톤",),
                (),
            ),
        ),
        unresolved_aliases=(UnresolvedDestinationAlias("미등록", "원본!9행"),),
        transport_subtotals=(
            TransportSubtotalInput(99, "미등록 ID", 2, 1, "원본!소계"),
        ),
    )

    first = validate_report(context)
    second = validate_report(context)

    assert first.issues == second.issues
    assert [(item.destination_id, item.code) for item in first.issues] == [
        (20, "MISSING_NEXT_MONTH_PLAN"),
        (20, "MISSING_VEHICLE_RATE"),
        (10, "MISSING_ACTUAL_QUANTITY"),
        (99, "TRANSPORT_SUBTOTAL_MISMATCH"),
        (None, "MISSING_DESTINATION_ALIAS"),
    ]


def test_issue_sorting_puts_errors_before_warnings():
    from app.domain.validation import ValidationIssue, _ordered_unique_issues

    warning = ValidationIssue(
        code="OPTIONAL_REVIEW_NOTE",
        severity="warning",
        message="참고용 검토 메모입니다.",
        destination_id=1,
    )
    error = ValidationIssue(
        code="MISSING_SALES",
        severity="error",
        message="매출 입력을 확인하세요.",
    )

    assert _ordered_unique_issues([warning, error], {1: 1}) == (error, warning)


def test_output_models_are_immutable_and_issues_are_a_tuple():
    from dataclasses import FrozenInstanceError

    from app.domain.validation import validate_report

    result = validate_report(_context())

    assert isinstance(result.issues, tuple)
    with pytest.raises(FrozenInstanceError):
        result.issues = ()


@pytest.mark.parametrize(
    "factory",
    [
        lambda: _context(report_month="2026-13"),
        lambda: _bare_context(report_month="0000-01"),
        lambda: _invalid_destination(destination_id=True),
        lambda: _invalid_destination(destination_id=0),
        lambda: _invalid_destination(actual_quantity=1),
        lambda: _invalid_destination(next_month_plan_quantity=0),
        lambda: _invalid_sales(amount_won=True),
        lambda: _invalid_batch(batch_id=False),
        lambda: _invalid_reconciliation(
            value_kind="quantity", detail_total=1, grand_total=1
        ),
        lambda: _invalid_reconciliation(
            value_kind="money", detail_total=Decimal("1"), grand_total=1
        ),
        lambda: _context(unresolved_aliases=[]),
        lambda: _invalid_history_complete_type(),
    ],
)
def test_invalid_context_is_rejected(factory):
    with pytest.raises((TypeError, ValueError)):
        factory()


def _invalid_destination(**changes):
    from app.domain.validation import DestinationValidationInput

    values = {
        "destination_id": 1,
        "name": "정상 납품처",
        "display_order": 1,
        "required_for_report": True,
        "actual_quantity": Decimal("1"),
        "next_month_plan_quantity": Decimal("1"),
    }
    values.update(changes)
    destination = DestinationValidationInput(**values)
    return _context(destinations=(destination,))


def _invalid_sales(**changes):
    from app.domain.validation import SalesInput

    values = {
        "amount_won": 1,
        "confirmed_at": "2026-08-31",
        "source_locator": "월 매출 입력",
    }
    values.update(changes)
    return _context(sales=SalesInput(**values))


def _invalid_batch(**changes):
    from app.domain.validation import ImportBatchInput

    values = {
        "batch_id": 1,
        "report_month": "2026-08",
        "source_type": "transport",
        "file_sha256": "sha",
        "source_locator": "배치 #1",
    }
    values.update(changes)
    return _context(current_import_batches=(ImportBatchInput(**values),))


def _invalid_reconciliation(**changes):
    from app.domain.validation import TotalReconciliationInput

    values = {
        "metric_label": "실적 수량",
        "value_kind": "quantity",
        "detail_total": Decimal("1"),
        "grand_total": Decimal("1"),
        "source_locator": "총계",
    }
    values.update(changes)
    return _context(total_reconciliations=(TotalReconciliationInput(**values),))


def _bare_context(**changes):
    from app.domain.validation import ValidationContext

    values = {
        "report_month": "2026-08",
        "sales": None,
        "prior_year_history": None,
    }
    values.update(changes)
    return ValidationContext(**values)


def _invalid_history_complete_type():
    from app.domain.calculations import AverageResult

    months = tuple(f"2025-{month:02d}" for month in range(1, 13))
    return _context(
        prior_year_history=AverageResult(
            expected_months=months,
            missing_months=(),
            value=Decimal("1"),
            complete=1,
        )
    )
