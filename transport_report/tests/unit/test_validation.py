from decimal import Decimal

import pytest


SHA_A = "a" * 64
SHA_B = "b" * 64


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
    from app.domain.validation import (
        ReportMonthSources,
        SalesInput,
        ValidationContext,
    )

    values = {
        "report_month": "2026-08",
        "sales": SalesInput(
            amount_won=1_000_000,
            confirmed_at="2026-08-31",
            source_locator="월 매출 입력",
        ),
        "prior_year_history": _complete_prior_year(),
        "month_sources": ReportMonthSources(
            plan_month="2026-08",
            actual_month="2026-08",
            transport_import_month="2026-08",
            ppt_title_month="2026-08",
            graph_last_month="2026-08",
        ),
    }
    values.update(changes)
    return ValidationContext(**values)


def _next_plan(quantity: Decimal | None, month: str | None = "2026-09"):
    from app.domain.validation import NextMonthPlanInput

    return NextMonthPlanInput(month, quantity)


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
                next_month_plan=_next_plan(Decimal("11")),
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
                next_month_plan=_next_plan(Decimal("1")),
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
                next_month_plan=_next_plan(Decimal("0")),
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
                ImportBatchInput(21, "2026-08", "transport", SHA_A, "배치 #21", True),
                ImportBatchInput(22, "2026-08", "transport", SHA_B, "배치 #22", True),
            )
        )
    )

    issue = next(issue for issue in result.issues if issue.code == "DUPLICATE_IMPORT")
    assert issue.severity == "error"
    assert issue.source_locator == "배치 #21, 배치 #22"
    assert "서로 다른" in issue.message
    assert "하나만" in issue.message
    assert "transport" in issue.message


def test_same_file_import_retry_is_idempotent_not_a_duplicate_conflict():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _context(
            current_import_batches=(
                ImportBatchInput(21, "2026-08", "transport", SHA_A, "배치 #21", True),
                ImportBatchInput(22, "2026-08", "transport", SHA_A, "재시도", True),
            )
        )
    )

    assert "DUPLICATE_IMPORT" not in {issue.code for issue in result.issues}


def test_current_batches_from_another_report_month_are_not_duplicate_candidates():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _context(
            current_import_batches=(
                ImportBatchInput(31, "2026-07", "transport", SHA_A, "배치 #31", True),
                ImportBatchInput(32, "2026-07", "transport", SHA_B, "배치 #32", True),
            )
        )
    )

    assert "DUPLICATE_IMPORT" not in {issue.code for issue in result.issues}


def test_duplicate_batch_locator_order_is_stable_when_batch_ids_tie():
    from app.domain.validation import ImportBatchInput, validate_report

    batches = (
        ImportBatchInput(21, "2026-08", "transport", SHA_B, "배치 Z", True),
        ImportBatchInput(21, "2026-08", "transport", SHA_A, "배치 A", True),
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
                    next_month_plan=None,
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
    ("field", "source_locator"),
    [
        ("plan_month", "당월 계획 입력"),
        ("actual_month", "당월 실적 입력"),
        ("transport_import_month", "운반비 가져오기"),
        ("ppt_title_month", "PPT 표지 제목"),
        ("graph_last_month", "그래프 마지막 월"),
    ],
)
def test_every_report_month_source_must_match_configured_month(field, source_locator):
    from dataclasses import replace

    from app.domain.validation import validate_report

    context = _context()
    wrong_sources = replace(context.month_sources, **{field: "2026-07"})
    result = validate_report(replace(context, month_sources=wrong_sources))

    issue = next(
        issue for issue in result.issues if issue.code == "REPORT_MONTH_MISMATCH"
    )
    assert issue.severity == "error"
    assert issue.source_locator == source_locator
    assert "2026-07" in issue.message
    assert "2026-08" in issue.message
    assert "수정" in issue.message


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
                _next_plan(Decimal("1")),
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
        lambda: _invalid_next_plan_quantity(0),
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
        "next_month_plan": _next_plan(Decimal("1")),
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
        "file_sha256": SHA_A,
        "source_locator": "배치 #1",
        "is_current": True,
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
    from app.domain.validation import ReportMonthSources, ValidationContext

    values = {
        "report_month": "2026-08",
        "sales": None,
        "prior_year_history": None,
        "month_sources": ReportMonthSources(None, None, None, None, None),
    }
    values.update(changes)
    return ValidationContext(**values)


def _invalid_next_plan_quantity(value):
    from app.domain.validation import NextMonthPlanInput

    return _invalid_destination(
        next_month_plan=NextMonthPlanInput("2026-09", value)
    )


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


@pytest.mark.parametrize(
    ("report_month", "plan_month"),
    [("2026-11", "2026-12"), ("2026-12", "2027-01")],
)
def test_zero_next_month_plan_is_present_across_year_rollover(
    report_month, plan_month
):
    from app.domain.validation import (
        DestinationValidationInput,
        NextMonthPlanInput,
        validate_report,
    )

    destination = DestinationValidationInput(
        destination_id=1,
        name="울산공장",
        display_order=1,
        required_for_report=True,
        actual_quantity=Decimal("0"),
        next_month_plan=NextMonthPlanInput(plan_month, Decimal("0")),
    )

    result = validate_report(
        _provenance_context(report_month=report_month, destinations=(destination,))
    )

    assert "MISSING_NEXT_MONTH_PLAN" not in {issue.code for issue in result.issues}
    assert "REPORT_MONTH_MISMATCH" not in {issue.code for issue in result.issues}


@pytest.mark.parametrize(
    "next_month_plan",
    [
        None,
        pytest.param((None, Decimal("0")), id="missing-month"),
        pytest.param(("2027-01", None), id="missing-quantity"),
    ],
)
def test_missing_next_month_plan_month_or_quantity_is_blocking(next_month_plan):
    from app.domain.validation import (
        DestinationValidationInput,
        NextMonthPlanInput,
        validate_report,
    )

    snapshot = (
        None
        if next_month_plan is None
        else NextMonthPlanInput(*next_month_plan)
    )
    destination = DestinationValidationInput(
        1,
        "울산공장",
        1,
        True,
        Decimal("1"),
        snapshot,
    )

    result = validate_report(
        _provenance_context(report_month="2026-12", destinations=(destination,))
    )

    assert "MISSING_NEXT_MONTH_PLAN" in {issue.code for issue in result.issues}


def test_wrong_next_month_plan_month_reports_both_month_and_missing_plan_errors():
    from app.domain.validation import (
        DestinationValidationInput,
        NextMonthPlanInput,
        validate_report,
    )

    destination = DestinationValidationInput(
        1,
        "울산공장",
        1,
        True,
        Decimal("1"),
        NextMonthPlanInput("2026-12", Decimal("0")),
    )

    result = validate_report(
        _provenance_context(report_month="2026-12", destinations=(destination,))
    )

    codes = {issue.code for issue in result.issues}
    assert {"MISSING_NEXT_MONTH_PLAN", "REPORT_MONTH_MISMATCH"} <= codes
    month_issue = next(
        issue for issue in result.issues if issue.code == "REPORT_MONTH_MISMATCH"
    )
    assert month_issue.source_locator == "다음 달 계획 입력"
    assert "2026-12" in month_issue.message
    assert "2027-01" in month_issue.message


def test_plan_month_mismatches_keep_destination_identity_order_and_dedup():
    from dataclasses import replace

    from app.domain.validation import (
        DestinationValidationInput,
        NextMonthPlanInput,
        validate_report,
    )

    wrong_plan = NextMonthPlanInput("2026-08", Decimal("1"))
    destinations = (
        DestinationValidationInput(
            10, "동일 공장명", 2, True, Decimal("1"), wrong_plan
        ),
        DestinationValidationInput(
            20, "동일 공장명", 1, True, Decimal("1"), wrong_plan
        ),
    )
    context = _provenance_context(destinations=destinations)
    global_mismatch = replace(context.month_sources, plan_month="2026-07")

    result = validate_report(replace(context, month_sources=global_mismatch))

    destination_issues = [
        (issue.destination_id, issue.code)
        for issue in result.issues
        if issue.destination_id is not None
        and issue.code in {"MISSING_NEXT_MONTH_PLAN", "REPORT_MONTH_MISMATCH"}
    ]
    assert destination_issues == [
        (20, "MISSING_NEXT_MONTH_PLAN"),
        (20, "REPORT_MONTH_MISMATCH"),
        (10, "MISSING_NEXT_MONTH_PLAN"),
        (10, "REPORT_MONTH_MISMATCH"),
    ]
    global_issue = next(
        issue
        for issue in result.issues
        if issue.code == "REPORT_MONTH_MISMATCH"
        and issue.source_locator == "당월 계획 입력"
    )
    assert global_issue.destination_id is None


def test_all_five_report_month_sources_matching_produces_no_month_issue():
    from app.domain.validation import validate_report

    result = validate_report(_provenance_context())

    assert "REPORT_MONTH_MISMATCH" not in {issue.code for issue in result.issues}


def test_report_month_source_rejects_malformed_nonmissing_month():
    from app.domain.validation import ReportMonthSources

    with pytest.raises(ValueError, match="ASCII YYYY-MM"):
        ReportMonthSources(
            plan_month="2026-13",
            actual_month="2026-08",
            transport_import_month="2026-08",
            ppt_title_month="2026-08",
            graph_last_month="2026-08",
        )


@pytest.mark.parametrize(
    ("field", "locator"),
    [
        ("plan_month", "당월 계획 입력"),
        ("actual_month", "당월 실적 입력"),
        ("transport_import_month", "운반비 가져오기"),
        ("ppt_title_month", "PPT 표지 제목"),
        ("graph_last_month", "그래프 마지막 월"),
    ],
)
def test_each_missing_report_month_source_is_reported(field, locator):
    from dataclasses import replace

    from app.domain.validation import validate_report

    context = _provenance_context()
    missing = replace(context.month_sources, **{field: None})
    result = validate_report(replace(context, month_sources=missing))

    issue = next(
        issue for issue in result.issues if issue.code == "REPORT_MONTH_MISMATCH"
    )
    assert issue.source_locator == locator
    assert "없습니다" in issue.message
    assert "2026-08" in issue.message


def test_superseded_import_batch_does_not_conflict_with_corrected_current_batch():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _provenance_context(
            current_import_batches=(
                ImportBatchInput(
                    1, "2026-08", "transport", SHA_A, "배치 #1", False
                ),
                ImportBatchInput(
                    2, "2026-08", "transport", SHA_B, "배치 #2", True
                ),
            )
        )
    )

    assert "DUPLICATE_IMPORT" not in {issue.code for issue in result.issues}


def test_two_distinct_current_import_batches_block_and_identify_the_source():
    from app.domain.validation import ImportBatchInput, validate_report

    result = validate_report(
        _provenance_context(
            current_import_batches=(
                ImportBatchInput(
                    1, "2026-08", "transport", SHA_A, "배치 #1", True
                ),
                ImportBatchInput(
                    2, "2026-08", "transport", SHA_B, "배치 #2", True
                ),
            )
        )
    )

    issue = next(issue for issue in result.issues if issue.code == "DUPLICATE_IMPORT")
    assert issue.source_locator == "배치 #1, 배치 #2"
    assert "transport" in issue.message
    assert "배치 #1" in issue.message
    assert "배치 #2" in issue.message


@pytest.mark.parametrize(
    "file_sha256",
    ["a" * 63, "g" * 64, "A" * 64],
    ids=("short", "non-hex", "uppercase"),
)
def test_import_batch_requires_lowercase_64_character_sha256(file_sha256):
    from app.domain.validation import ImportBatchInput

    with pytest.raises(ValueError, match="64 lowercase hexadecimal"):
        ImportBatchInput(
            1, "2026-08", "transport", file_sha256, "배치 #1", True
        )


def test_import_batch_current_state_is_strict_bool():
    from app.domain.validation import ImportBatchInput

    with pytest.raises(TypeError, match="is_current"):
        ImportBatchInput(1, "2026-08", "transport", SHA_A, "배치 #1", 1)


def _provenance_context(**changes):
    from app.domain.validation import (
        ReportMonthSources,
        SalesInput,
        ValidationContext,
    )

    report_month = changes.get("report_month", "2026-08")
    values = {
        "report_month": report_month,
        "sales": SalesInput(1_000_000, f"{report_month}-28", "월 매출 입력"),
        "prior_year_history": _complete_prior_year(report_month),
        "month_sources": ReportMonthSources(
            plan_month=report_month,
            actual_month=report_month,
            transport_import_month=report_month,
            ppt_title_month=report_month,
            graph_last_month=report_month,
        ),
    }
    values.update(changes)
    return ValidationContext(**values)
