"""Pure preflight validation for monthly report generation."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Literal

from app.domain.calculations import AverageResult


Severity = Literal["error", "warning"]
ReconciliationKind = Literal["quantity", "money"]

_MONTH = re.compile(r"^[0-9]{4}-(0[1-9]|1[0-2])$")
_SHA256 = re.compile(r"^[0-9a-fA-F]{64}$")
_MAX_SQLITE_INTEGER = 2**63 - 1
_MONTH_SOURCE_FIELDS = (
    ("plan_month", "당월 계획", "당월 계획 입력"),
    ("actual_month", "당월 실적", "당월 실적 입력"),
    ("transport_import_month", "운반비 가져오기 자료", "운반비 가져오기"),
    ("ppt_title_month", "PPT 표지 제목", "PPT 표지 제목"),
    ("graph_last_month", "그래프 마지막 월", "그래프 마지막 월"),
)
_BLOCKING_CODES = frozenset(
    {
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
    }
)


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    code: str
    severity: Severity
    message: str
    destination_id: int | None = None
    source_locator: str | None = None

    def __post_init__(self) -> None:
        _require_text(self.code, "code")
        if self.severity not in {"error", "warning"}:
            raise ValueError("severity must be error or warning")
        if self.code in _BLOCKING_CODES and self.severity != "error":
            raise ValueError(f"{self.code} is a blocking validation code")
        _require_text(self.message, "message")
        if self.destination_id is not None:
            _require_positive_id(self.destination_id, "destination_id")
        if self.source_locator is not None:
            _require_text(self.source_locator, "source_locator")


@dataclass(frozen=True, slots=True)
class UnresolvedDestinationAlias:
    alias: str
    source_locator: str

    def __post_init__(self) -> None:
        _require_text(self.alias, "alias")
        _require_text(self.source_locator, "source_locator")


@dataclass(frozen=True, slots=True)
class SalesInput:
    amount_won: int | None
    confirmed_at: str | None
    source_locator: str

    def __post_init__(self) -> None:
        if self.amount_won is not None:
            _require_nonnegative_int(self.amount_won, "amount_won")
        if self.confirmed_at is not None:
            _require_text(self.confirmed_at, "confirmed_at")
        _require_text(self.source_locator, "source_locator")


@dataclass(frozen=True, slots=True)
class NextMonthPlanInput:
    report_month: str | None
    quantity: Decimal | None

    def __post_init__(self) -> None:
        if self.report_month is not None:
            _require_month(self.report_month, "next_month_plan.report_month")
        _require_optional_quantity(self.quantity, "next_month_plan.quantity")


@dataclass(frozen=True, slots=True)
class DestinationValidationInput:
    destination_id: int
    name: str
    display_order: int
    required_for_report: bool
    actual_quantity: Decimal | None
    next_month_plan: NextMonthPlanInput | None
    required_vehicle_types: tuple[str, ...] = ()
    rated_vehicle_types: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        _require_positive_id(self.destination_id, "destination_id")
        _require_text(self.name, "name")
        _require_nonnegative_int(self.display_order, "display_order")
        if not isinstance(self.required_for_report, bool):
            raise TypeError("required_for_report must be bool")
        _require_optional_quantity(self.actual_quantity, "actual_quantity")
        if self.next_month_plan is not None and not isinstance(
            self.next_month_plan, NextMonthPlanInput
        ):
            raise TypeError("next_month_plan must be NextMonthPlanInput or None")
        _require_text_tuple(self.required_vehicle_types, "required_vehicle_types")
        _require_text_tuple(self.rated_vehicle_types, "rated_vehicle_types")


@dataclass(frozen=True, slots=True)
class ImportBatchInput:
    batch_id: int
    report_month: str
    source_type: str
    file_sha256: str
    source_locator: str
    is_current: bool

    def __post_init__(self) -> None:
        _require_positive_id(self.batch_id, "batch_id")
        _require_month(self.report_month, "report_month")
        _require_text(self.source_type, "source_type")
        if not isinstance(self.file_sha256, str) or _SHA256.fullmatch(
            self.file_sha256
        ) is None:
            raise ValueError("file_sha256 must be 64 hexadecimal characters")
        object.__setattr__(self, "file_sha256", self.file_sha256.lower())
        _require_text(self.source_locator, "source_locator")
        if not isinstance(self.is_current, bool):
            raise TypeError("is_current must be bool")


@dataclass(frozen=True, slots=True)
class TransportSubtotalInput:
    destination_id: int
    destination_name: str
    detail_total_won: int
    reported_subtotal_won: int
    source_locator: str

    def __post_init__(self) -> None:
        _require_positive_id(self.destination_id, "destination_id")
        _require_text(self.destination_name, "destination_name")
        _require_nonnegative_int(self.detail_total_won, "detail_total_won")
        _require_nonnegative_int(
            self.reported_subtotal_won, "reported_subtotal_won"
        )
        _require_text(self.source_locator, "source_locator")


@dataclass(frozen=True, slots=True)
class TotalReconciliationInput:
    metric_label: str
    value_kind: ReconciliationKind
    detail_total: Decimal | int
    grand_total: Decimal | int
    source_locator: str

    def __post_init__(self) -> None:
        _require_text(self.metric_label, "metric_label")
        if self.value_kind == "quantity":
            _require_quantity(self.detail_total, "detail_total")
            _require_quantity(self.grand_total, "grand_total")
        elif self.value_kind == "money":
            _require_nonnegative_int(self.detail_total, "detail_total")
            _require_nonnegative_int(self.grand_total, "grand_total")
        else:
            raise ValueError("value_kind must be quantity or money")
        _require_text(self.source_locator, "source_locator")


@dataclass(frozen=True, slots=True)
class ReportMonthSources:
    plan_month: str | None
    actual_month: str | None
    transport_import_month: str | None
    ppt_title_month: str | None
    graph_last_month: str | None

    def __post_init__(self) -> None:
        for field, _, _ in _MONTH_SOURCE_FIELDS:
            value = getattr(self, field)
            if value is not None:
                _require_month(value, field)


@dataclass(frozen=True, slots=True)
class ValidationContext:
    report_month: str
    sales: SalesInput | None
    prior_year_history: AverageResult | None
    month_sources: ReportMonthSources
    unresolved_aliases: tuple[UnresolvedDestinationAlias, ...] = ()
    destinations: tuple[DestinationValidationInput, ...] = ()
    current_import_batches: tuple[ImportBatchInput, ...] = ()
    transport_subtotals: tuple[TransportSubtotalInput, ...] = ()
    total_reconciliations: tuple[TotalReconciliationInput, ...] = ()

    def __post_init__(self) -> None:
        _require_month(self.report_month, "report_month")
        if self.sales is not None and not isinstance(self.sales, SalesInput):
            raise TypeError("sales must be SalesInput or None")
        _require_history(self.prior_year_history, self.report_month)
        if not isinstance(self.month_sources, ReportMonthSources):
            raise TypeError("month_sources must be ReportMonthSources")
        _require_tuple_items(
            self.unresolved_aliases,
            UnresolvedDestinationAlias,
            "unresolved_aliases",
        )
        _require_tuple_items(
            self.destinations, DestinationValidationInput, "destinations"
        )
        _require_tuple_items(
            self.current_import_batches,
            ImportBatchInput,
            "current_import_batches",
        )
        _require_tuple_items(
            self.transport_subtotals,
            TransportSubtotalInput,
            "transport_subtotals",
        )
        _require_tuple_items(
            self.total_reconciliations,
            TotalReconciliationInput,
            "total_reconciliations",
        )
        destination_ids = [item.destination_id for item in self.destinations]
        if len(destination_ids) != len(set(destination_ids)):
            raise ValueError("destinations contain a duplicate destination_id")


@dataclass(frozen=True, slots=True)
class ValidationResult:
    issues: tuple[ValidationIssue, ...]

    def __post_init__(self) -> None:
        _require_tuple_items(self.issues, ValidationIssue, "issues")

    @property
    def can_generate(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)


def validate_report(context: ValidationContext) -> ValidationResult:
    """Return ordered, de-duplicated preflight issues for one report month.

    Import-batch checks validate which batch is selected for the report. The
    importer's unique source identity and reconciliation continue to prevent
    row duplication; this validator does not duplicate those responsibilities.
    """
    if not isinstance(context, ValidationContext):
        raise TypeError("context must be ValidationContext")

    issues: list[ValidationIssue] = []
    for item in context.unresolved_aliases:
        issues.append(
            ValidationIssue(
                code="MISSING_DESTINATION_ALIAS",
                severity="error",
                message=(
                    f"납품처 '{item.alias}'의 별칭을 찾을 수 없습니다. "
                    "기준정보의 납품처 별칭에 등록한 뒤 "
                    f"{item.source_locator}을 다시 확인하세요."
                ),
                source_locator=item.source_locator,
            )
        )

    missing_history = (
        _prior_year_months(context.report_month)
        if context.prior_year_history is None
        else context.prior_year_history.missing_months
    )
    if missing_history:
        issues.append(
            ValidationIssue(
                code="MISSING_PRIOR_YEAR_HISTORY",
                severity="error",
                message=(
                    "전년도 전체 이력 중 "
                    f"{', '.join(missing_history)} 자료가 없습니다. "
                    "전년도 비교 자료를 보완하세요."
                ),
                source_locator="전년도 이력",
            )
        )

    for field, source_label, source_locator in _MONTH_SOURCE_FIELDS:
        source_month = getattr(context.month_sources, field)
        if source_month != context.report_month:
            issues.append(
                _month_mismatch_issue(
                    source_label=source_label,
                    source_month=source_month,
                    report_month=context.report_month,
                    source_locator=source_locator,
                )
            )

    if (
        context.sales is None
        or context.sales.amount_won is None
        or context.sales.confirmed_at is None
    ):
        source_locator = (
            "월 매출 입력" if context.sales is None else context.sales.source_locator
        )
        issues.append(
            ValidationIssue(
                code="MISSING_SALES",
                severity="error",
                message=(
                    "당월 매출액 또는 확정일이 없습니다. "
                    "월 매출 입력에서 매출액과 확정일을 모두 입력하세요."
                ),
                source_locator=source_locator,
            )
        )

    batches_by_source: dict[str, list[ImportBatchInput]] = {}
    for batch in context.current_import_batches:
        if batch.is_current and batch.report_month != context.report_month:
            issues.append(
                _month_mismatch_issue(
                    source_label=(
                        f"{batch.source_type} 운반비 가져오기 배치 #{batch.batch_id}"
                    ),
                    source_month=batch.report_month,
                    report_month=context.report_month,
                    source_locator=batch.source_locator,
                )
            )
        if batch.is_current and batch.report_month == context.report_month:
            batches_by_source.setdefault(batch.source_type, []).append(batch)
    for source_type, batches in batches_by_source.items():
        signatures = {batch.file_sha256 for batch in batches}
        if len(signatures) > 1:
            locators = ", ".join(
                dict.fromkeys(
                    batch.source_locator
                    for batch in sorted(
                        batches,
                        key=lambda item: (
                            item.batch_id,
                            item.source_locator,
                            item.file_sha256,
                        ),
                    )
                )
            )
            issues.append(
                ValidationIssue(
                    code="DUPLICATE_IMPORT",
                    severity="error",
                    message=(
                        f"원천 '{source_type}'에 서로 다른 현재 운반비 가져오기 배치"
                        f"({locators})가 있습니다. "
                        "사용할 배치 하나만 남기고 다시 생성하세요."
                    ),
                    source_locator=locators,
                )
            )

    for subtotal in context.transport_subtotals:
        if subtotal.detail_total_won != subtotal.reported_subtotal_won:
            issues.append(
                ValidationIssue(
                    code="TRANSPORT_SUBTOTAL_MISMATCH",
                    severity="error",
                    message=(
                        f"{subtotal.destination_name} 운반비 상세 행 합계 "
                        f"{subtotal.detail_total_won:,}원과 소계 "
                        f"{subtotal.reported_subtotal_won:,}원이 다릅니다. "
                        f"{subtotal.source_locator}의 상세 행과 소계를 수정하세요."
                    ),
                    destination_id=subtotal.destination_id,
                    source_locator=subtotal.source_locator,
                )
            )

    for reconciliation in context.total_reconciliations:
        if reconciliation.detail_total != reconciliation.grand_total:
            issues.append(
                ValidationIssue(
                    code="TOTAL_RECONCILIATION_FAILED",
                    severity="error",
                    message=(
                        f"{reconciliation.metric_label}의 납품처/상세 합계 "
                        f"{reconciliation.detail_total}와 보고서 총계 "
                        f"{reconciliation.grand_total}가 일치하지 않습니다. "
                        "포함 기준과 원본 상세 값을 확인해 총계를 맞추세요."
                    ),
                    source_locator=reconciliation.source_locator,
                )
            )

    for destination in context.destinations:
        if destination.required_for_report and destination.actual_quantity is None:
            issues.append(
                ValidationIssue(
                    code="MISSING_ACTUAL_QUANTITY",
                    severity="error",
                    message=(
                        f"{destination.name}의 당월 실적 수량이 없습니다. "
                        "당월 실적 입력에서 수량을 입력하세요(실적 없음은 0 입력)."
                    ),
                    destination_id=destination.destination_id,
                    source_locator="당월 실적 입력",
                )
            )
        expected_plan_month = _next_month(context.report_month)
        next_plan = destination.next_month_plan
        has_expected_plan = (
            next_plan is not None
            and next_plan.report_month == expected_plan_month
            and next_plan.quantity is not None
        )
        if destination.required_for_report and not has_expected_plan:
            issues.append(
                ValidationIssue(
                    code="MISSING_NEXT_MONTH_PLAN",
                    severity="error",
                    message=(
                        f"{destination.name}의 {_next_month(context.report_month)} "
                        "계획이 없습니다. 다음 달 계획 입력에서 수량과 비용을 입력하세요."
                    ),
                    destination_id=destination.destination_id,
                    source_locator="다음 달 계획 입력",
                )
            )
        if (
            destination.required_for_report
            and next_plan is not None
            and next_plan.report_month is not None
            and next_plan.report_month != expected_plan_month
        ):
            issues.append(
                _month_mismatch_issue(
                    source_label=f"{destination.name} 다음 달 계획",
                    source_month=next_plan.report_month,
                    report_month=expected_plan_month,
                    source_locator="다음 달 계획 입력",
                    expected_label="예상 계획 월",
                    destination_id=destination.destination_id,
                )
            )
        for vehicle_type in sorted(
            set(destination.required_vehicle_types)
            - set(destination.rated_vehicle_types)
        ):
            issues.append(
                ValidationIssue(
                    code="MISSING_VEHICLE_RATE",
                    severity="error",
                    message=(
                        f"{destination.name}의 {vehicle_type} 차량 요율이 없습니다. "
                        "차량 요율 기준정보에 해당 월의 요율을 등록하세요."
                    ),
                    destination_id=destination.destination_id,
                    source_locator="차량 요율 기준정보",
                )
            )

    display_orders = {
        item.destination_id: item.display_order for item in context.destinations
    }
    return ValidationResult(_ordered_unique_issues(issues, display_orders))


def _month_mismatch_issue(
    *,
    source_label: str,
    source_month: str | None,
    report_month: str,
    source_locator: str,
    expected_label: str = "설정된 보고 월",
    destination_id: int | None = None,
) -> ValidationIssue:
    if source_month is None:
        message = (
            f"{source_label}의 기준 월이 없습니다. {source_locator}에서 "
            f"{expected_label} {report_month}을 입력하세요."
        )
    else:
        message = (
            f"{source_label}의 기준 월 {source_month}이 {expected_label} "
            f"{report_month}과 다릅니다. {source_locator}의 월을 수정하세요."
        )
    return ValidationIssue(
        code="REPORT_MONTH_MISMATCH",
        severity="error",
        message=message,
        destination_id=destination_id,
        source_locator=source_locator,
    )


def _ordered_unique_issues(
    issues: list[ValidationIssue], display_orders: dict[int, int]
) -> tuple[ValidationIssue, ...]:
    def sort_key(issue: ValidationIssue) -> tuple[object, ...]:
        if issue.destination_id in display_orders:
            destination_key = (
                0,
                display_orders[issue.destination_id],
                issue.destination_id,
            )
        elif issue.destination_id is not None:
            destination_key = (1, issue.destination_id, issue.destination_id)
        else:
            destination_key = (2, 0, 0)
        return (
            0 if issue.severity == "error" else 1,
            *destination_key,
            issue.code,
            issue.source_locator or "",
            issue.message,
        )

    return tuple(sorted(set(issues), key=sort_key))


def _prior_year_months(report_month: str) -> tuple[str, ...]:
    year = int(report_month[:4]) - 1
    return tuple(f"{year:04d}-{month:02d}" for month in range(1, 13))


def _next_month(report_month: str) -> str:
    current = date.fromisoformat(f"{report_month}-01")
    if current.month == 12:
        return f"{current.year + 1:04d}-01"
    return f"{current.year:04d}-{current.month + 1:02d}"


def _require_history(value: AverageResult | None, report_month: str) -> None:
    if value is None:
        return
    if not isinstance(value, AverageResult):
        raise TypeError("prior_year_history must be AverageResult or None")
    _require_text_tuple(value.expected_months, "prior_year_history.expected_months")
    _require_text_tuple(value.missing_months, "prior_year_history.missing_months")
    expected = _prior_year_months(report_month)
    if value.expected_months != expected:
        raise ValueError("prior_year_history must cover the previous calendar year")
    missing = tuple(month for month in expected if month in value.missing_months)
    if value.missing_months != missing:
        raise ValueError("prior_year_history.missing_months are invalid")
    if not isinstance(value.complete, bool):
        raise TypeError("prior_year_history.complete must be bool")
    if value.complete != (not value.missing_months):
        raise ValueError("prior_year_history completeness is inconsistent")
    if value.complete and value.value is None:
        raise ValueError("complete prior_year_history requires a value")
    if not value.complete and value.value is not None:
        raise ValueError("incomplete prior_year_history cannot have a value")
    if value.value is not None:
        _require_quantity(value.value, "prior_year_history.value")


def _require_tuple_items(value: object, item_type: type, field: str) -> None:
    if not isinstance(value, tuple):
        raise TypeError(f"{field} must be a tuple")
    if any(not isinstance(item, item_type) for item in value):
        raise TypeError(f"{field} contains an invalid item")


def _require_text_tuple(value: object, field: str) -> None:
    if not isinstance(value, tuple):
        raise TypeError(f"{field} must be a tuple")
    for item in value:
        _require_text(item, field)
    if len(value) != len(set(value)):
        raise ValueError(f"{field} contains duplicates")


def _require_month(value: object, field: str) -> None:
    if not isinstance(value, str) or _MONTH.fullmatch(value) is None:
        raise ValueError(f"{field} must use ASCII YYYY-MM")
    try:
        date.fromisoformat(f"{value}-01")
    except ValueError as error:
        raise ValueError(f"{field} must use a valid ASCII YYYY-MM") from error


def _require_text(value: object, field: str) -> None:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} must be nonblank text")


def _require_positive_id(value: object, field: str) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < 1
        or value > _MAX_SQLITE_INTEGER
    ):
        raise ValueError(f"{field} must be a positive integer")


def _require_nonnegative_int(value: object, field: str) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < 0
        or value > _MAX_SQLITE_INTEGER
    ):
        raise ValueError(f"{field} must be a nonnegative integer")


def _require_optional_quantity(value: object, field: str) -> None:
    if value is not None:
        _require_quantity(value, field)


def _require_quantity(value: object, field: str) -> None:
    if not isinstance(value, Decimal) or not value.is_finite() or value < 0:
        raise ValueError(f"{field} must be a nonnegative finite Decimal")
