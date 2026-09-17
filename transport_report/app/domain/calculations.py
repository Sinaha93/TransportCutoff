"""Pure calculations for monthly transport reporting.

Unit costs use :data:`decimal.ROUND_HALF_UP` because Excel's ``ROUND``
function rounds a positive half away from zero. Quantities and costs in this
domain are nonnegative, so this reproduces the workbook's two-decimal unit-cost
presentation. Other Decimal results are kept unrounded for downstream output.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum

from app.domain.models import Destination, GroupMember


UNIT_COST_QUANTUM = Decimal("0.01")
DEFAULT_REVIEW_THRESHOLD = Decimal("0.15")
MAX_SQLITE_INTEGER = 2**63 - 1


class CalculationKind(Enum):
    DESTINATION = "destination"
    DERIVED_GROUP = "derived_group"
    GRAND_TOTAL = "grand_total"


@dataclass(frozen=True, slots=True)
class DestinationCalculation:
    planned_quantity: Decimal | None
    planned_cost_won: int | None
    actual_quantity: Decimal | None
    actual_cost_won: int | None
    planned_unit_cost: Decimal | None
    actual_unit_cost: Decimal | None
    quantity_variance: Decimal | None
    quantity_variance_pct: Decimal | None
    cost_variance_won: int | None
    cost_variance_pct: Decimal | None
    actual_unit_cost_variance: Decimal | None
    actual_unit_cost_variance_pct: Decimal | None
    kind: CalculationKind = CalculationKind.DESTINATION

    def __post_init__(self) -> None:
        if not isinstance(self.kind, CalculationKind):
            raise ValueError("kind must be a CalculationKind")
        if self.kind is CalculationKind.DESTINATION:
            _validate_period_inputs(
                self.planned_quantity, self.planned_cost_won, "planned"
            )
            _validate_period_inputs(
                self.actual_quantity, self.actual_cost_won, "actual"
            )
        else:
            _require_optional_nonnegative_decimal(
                self.planned_quantity, "planned_quantity"
            )
            _require_optional_nonnegative_decimal(
                self.actual_quantity, "actual_quantity"
            )
            _require_optional_nonnegative_int(
                self.planned_cost_won, "planned_cost_won"
            )
            _require_optional_nonnegative_int(self.actual_cost_won, "actual_cost_won")
        _require_optional_nonnegative_decimal(
            self.planned_unit_cost, "planned_unit_cost"
        )
        _require_optional_nonnegative_decimal(
            self.actual_unit_cost, "actual_unit_cost"
        )
        _require_optional_finite_decimal(self.quantity_variance, "quantity_variance")
        _require_optional_finite_decimal(
            self.quantity_variance_pct, "quantity_variance_pct"
        )
        _require_optional_int(self.cost_variance_won, "cost_variance_won")
        _require_optional_finite_decimal(self.cost_variance_pct, "cost_variance_pct")
        _require_optional_finite_decimal(
            self.actual_unit_cost_variance, "actual_unit_cost_variance"
        )
        _require_optional_finite_decimal(
            self.actual_unit_cost_variance_pct, "actual_unit_cost_variance_pct"
        )


@dataclass(frozen=True, slots=True)
class AverageResult:
    """An average and the calendar-month completeness behind it."""

    expected_months: tuple[str, ...]
    missing_months: tuple[str, ...]
    value: Decimal | None
    complete: bool


@dataclass(frozen=True, slots=True)
class HistoricalAverages:
    three_month: AverageResult
    six_month: AverageResult
    twelve_month: AverageResult
    comparison_year: AverageResult


@dataclass(frozen=True, slots=True)
class ReviewCandidate:
    destination_id: int
    name: str
    display_order: int
    calculation: DestinationCalculation


@dataclass(frozen=True, slots=True)
class ReviewItem:
    destination_id: int
    name: str
    display_order: int
    variance_pct: Decimal


@dataclass(frozen=True, slots=True)
class ReviewValidationItem:
    destination_id: int
    name: str
    display_order: int
    code: str
    message: str


@dataclass(frozen=True, slots=True)
class ReviewSelection:
    automatic_items: tuple[ReviewItem, ...]
    validation_items: tuple[ReviewValidationItem, ...]


def unit_cost(cost_won: int, quantity_ea: Decimal) -> Decimal | None:
    """Return transport cost per delivered EA, rounded to two decimals."""
    _require_nonnegative_int(
        cost_won, "cost_won", maximum=MAX_SQLITE_INTEGER
    )
    _require_nonnegative_decimal(quantity_ea, "quantity_ea")
    return _rounded_unit_cost(cost_won, quantity_ea)


def _rounded_unit_cost(cost_won: int, quantity_ea: Decimal) -> Decimal | None:
    if quantity_ea == 0:
        return None
    return (Decimal(cost_won) / quantity_ea).quantize(
        UNIT_COST_QUANTUM, rounding=ROUND_HALF_UP
    )


def variance(actual: Decimal, plan: Decimal) -> Decimal:
    """Return the unrounded absolute variance ``actual - plan``."""
    _require_finite_decimal(actual, "actual")
    _require_finite_decimal(plan, "plan")
    return actual - plan


def variance_pct(actual: Decimal, plan: Decimal) -> Decimal | None:
    """Return the unrounded relative variance, or ``None`` for a zero plan."""
    _require_finite_decimal(actual, "actual")
    _require_finite_decimal(plan, "plan")
    return None if plan == 0 else variance(actual, plan) / plan


def calculate_destination(
    planned_quantity: Decimal | None,
    planned_cost_won: int | None,
    actual_quantity: Decimal | None,
    actual_cost_won: int | None,
) -> DestinationCalculation:
    """Calculate one destination while preserving a fully missing period.

    Quantity and cost for a period must either both be supplied or both be
    missing. This keeps partial source data from looking like a valid row.
    """
    _validate_period_inputs(planned_quantity, planned_cost_won, "planned")
    _validate_period_inputs(actual_quantity, actual_cost_won, "actual")
    return _make_calculation(
        planned_quantity,
        planned_cost_won,
        actual_quantity,
        actual_cost_won,
        kind=CalculationKind.DESTINATION,
    )


def calculate_total(
    calculations: Mapping[int, DestinationCalculation],
    destinations: Iterable[Destination],
) -> DestinationCalculation:
    """Aggregate direct destinations using their editable total flags.

    Derived report groups are not accepted as total rules, so calculations
    with keys that do not correspond to a supplied destination are ignored.
    Missing included values propagate as missing rather than becoming zero.
    """
    rules = tuple(destinations)
    _require_unique_ids((item.id for item in rules), "destinations")
    _validate_direct_sources(calculations, rules)
    return _make_calculation(
        _sum_metric(calculations, rules, "planned_quantity", "include_quantity_total", Decimal(0)),
        _sum_metric(calculations, rules, "planned_cost_won", "include_cost_total", 0),
        _sum_metric(calculations, rules, "actual_quantity", "include_quantity_total", Decimal(0)),
        _sum_metric(calculations, rules, "actual_cost_won", "include_cost_total", 0),
        kind=CalculationKind.GRAND_TOTAL,
    )


def calculate_group(
    calculations: Mapping[int, DestinationCalculation],
    members: Iterable[GroupMember],
) -> DestinationCalculation:
    """Aggregate a derived display group using its editable member flags."""
    rules = tuple(members)
    _require_unique_ids((item.destination_id for item in rules), "group members")
    _validate_direct_sources(calculations, rules)
    return _make_calculation(
        _sum_metric(calculations, rules, "planned_quantity", "include_quantity", Decimal(0)),
        _sum_metric(calculations, rules, "planned_cost_won", "include_cost", 0),
        _sum_metric(calculations, rules, "actual_quantity", "include_quantity", Decimal(0)),
        _sum_metric(calculations, rules, "actual_cost_won", "include_cost", 0),
        kind=CalculationKind.DERIVED_GROUP,
    )


def historical_averages(
    report_month: str, monthly_values: Mapping[str, Decimal | None]
) -> HistoricalAverages:
    """Calculate complete prior-month windows and the prior calendar year.

    A missing key and an explicit ``None`` both mark an incomplete period. No
    average is returned until every expected calendar month is present.
    """
    report_year, _ = _parse_month(report_month, "report_month")
    _validate_monthly_values(monthly_values)
    return HistoricalAverages(
        three_month=_average_for_months(monthly_values, _prior_months(report_month, 3)),
        six_month=_average_for_months(monthly_values, _prior_months(report_month, 6)),
        twelve_month=_average_for_months(monthly_values, _prior_months(report_month, 12)),
        comparison_year=_average_for_months(
            monthly_values,
            tuple(f"{report_year - 1:04d}-{month:02d}" for month in range(1, 13)),
        ),
    )


def select_unit_cost_reviews(
    candidates: Iterable[ReviewCandidate],
    threshold: Decimal = DEFAULT_REVIEW_THRESHOLD,
) -> ReviewSelection:
    """Select inclusive unit-cost variances and surface unavailable values."""
    _require_nonnegative_decimal(threshold, "threshold")
    candidates = tuple(candidates)
    for candidate in candidates:
        _validate_review_candidate(candidate)
    ordered = sorted(candidates, key=lambda item: (item.display_order, item.destination_id))
    _require_unique_ids((item.destination_id for item in ordered), "candidates")
    automatic: list[ReviewItem] = []
    validation: list[ReviewValidationItem] = []
    for candidate in ordered:
        value = candidate.calculation.actual_unit_cost_variance_pct
        if value is None:
            validation.append(
                ReviewValidationItem(
                    destination_id=candidate.destination_id,
                    name=candidate.name,
                    display_order=candidate.display_order,
                    code="UNIT_COST_VARIANCE_UNAVAILABLE",
                    message="Unit cost or its plan variance is unavailable.",
                )
            )
        elif value <= -threshold or value >= threshold:
            automatic.append(
                ReviewItem(
                    destination_id=candidate.destination_id,
                    name=candidate.name,
                    display_order=candidate.display_order,
                    variance_pct=value,
                )
            )
    return ReviewSelection(tuple(automatic), tuple(validation))


def _make_calculation(
    planned_quantity: Decimal | None,
    planned_cost_won: int | None,
    actual_quantity: Decimal | None,
    actual_cost_won: int | None,
    *,
    kind: CalculationKind,
) -> DestinationCalculation:
    planned_unit = _optional_unit_cost(planned_cost_won, planned_quantity)
    actual_unit = _optional_unit_cost(actual_cost_won, actual_quantity)
    quantity_difference = _optional_variance(actual_quantity, planned_quantity)
    cost_difference = (
        None
        if actual_cost_won is None or planned_cost_won is None
        else actual_cost_won - planned_cost_won
    )
    unit_difference = _optional_variance(actual_unit, planned_unit)
    return DestinationCalculation(
        planned_quantity=planned_quantity,
        planned_cost_won=planned_cost_won,
        actual_quantity=actual_quantity,
        actual_cost_won=actual_cost_won,
        planned_unit_cost=planned_unit,
        actual_unit_cost=actual_unit,
        quantity_variance=quantity_difference,
        quantity_variance_pct=_optional_variance_pct(actual_quantity, planned_quantity),
        cost_variance_won=cost_difference,
        cost_variance_pct=_optional_variance_pct(
            None if actual_cost_won is None else Decimal(actual_cost_won),
            None if planned_cost_won is None else Decimal(planned_cost_won),
        ),
        actual_unit_cost_variance=unit_difference,
        actual_unit_cost_variance_pct=_optional_variance_pct(actual_unit, planned_unit),
        kind=kind,
    )


def _optional_unit_cost(
    cost_won: int | None, quantity_ea: Decimal | None
) -> Decimal | None:
    if cost_won is None or quantity_ea is None:
        return None
    return _rounded_unit_cost(cost_won, quantity_ea)


def _optional_variance(
    actual: Decimal | None, plan: Decimal | None
) -> Decimal | None:
    if actual is None or plan is None:
        return None
    return variance(actual, plan)


def _optional_variance_pct(
    actual: Decimal | None, plan: Decimal | None
) -> Decimal | None:
    if actual is None or plan is None:
        return None
    return variance_pct(actual, plan)


def _sum_metric(
    calculations: Mapping[int, DestinationCalculation],
    rules: tuple[Destination, ...] | tuple[GroupMember, ...],
    value_field: str,
    include_field: str,
    zero: Decimal | int,
) -> Decimal | int | None:
    values: list[Decimal | int] = []
    for rule in rules:
        if not getattr(rule, include_field):
            continue
        destination_id = rule.id if isinstance(rule, Destination) else rule.destination_id
        calculation = calculations.get(destination_id)
        if calculation is None:
            return None
        if not isinstance(calculation, DestinationCalculation):
            raise ValueError("calculations must contain DestinationCalculation values")
        if calculation.kind is not CalculationKind.DESTINATION:
            raise ValueError(
                "totals and groups require direct destination calculations; "
                "a derived group or total cannot replace a destination"
            )
        value = getattr(calculation, value_field)
        if value is None:
            return None
        values.append(value)
    return sum(values, start=zero)


def _validate_direct_sources(
    calculations: Mapping[int, DestinationCalculation],
    rules: tuple[Destination, ...] | tuple[GroupMember, ...],
) -> None:
    for rule in rules:
        destination_id = rule.id if isinstance(rule, Destination) else rule.destination_id
        calculation = calculations.get(destination_id)
        if calculation is None:
            continue
        if not isinstance(calculation, DestinationCalculation):
            raise ValueError("calculations must contain DestinationCalculation values")
        if calculation.kind is not CalculationKind.DESTINATION:
            raise ValueError(
                "totals and groups require direct destination calculations; "
                "a derived group or total cannot replace a destination"
            )


def _average_for_months(
    monthly_values: Mapping[str, Decimal | None], expected_months: tuple[str, ...]
) -> AverageResult:
    missing = tuple(month for month in expected_months if monthly_values.get(month) is None)
    if missing:
        return AverageResult(expected_months, missing, None, False)
    total = sum((monthly_values[month] for month in expected_months), Decimal(0))
    return AverageResult(expected_months, (), total / Decimal(len(expected_months)), True)


def _prior_months(report_month: str, count: int) -> tuple[str, ...]:
    year, month = _parse_month(report_month, "report_month")
    ordinal = year * 12 + month - 1
    result: list[str] = []
    for offset in range(count, 0, -1):
        prior_ordinal = ordinal - offset
        prior_year, zero_based_month = divmod(prior_ordinal, 12)
        result.append(f"{prior_year:04d}-{zero_based_month + 1:02d}")
    return tuple(result)


def _parse_month(value: object, field: str) -> tuple[int, int]:
    if not isinstance(value, str) or len(value) != 7 or value[4] != "-":
        raise ValueError(f"{field} must be in YYYY-MM format")
    try:
        year = int(value[:4])
        month = int(value[5:])
    except ValueError as error:
        raise ValueError(f"{field} must be in YYYY-MM format") from error
    if year < 1 or not 1 <= month <= 12 or value != f"{year:04d}-{month:02d}":
        raise ValueError(f"{field} must be in YYYY-MM format")
    return year, month


def _validate_monthly_values(values: Mapping[str, Decimal | None]) -> None:
    if not isinstance(values, Mapping):
        raise ValueError("monthly_values must be a mapping")
    for month, value in values.items():
        _parse_month(month, "monthly_values month")
        if value is not None:
            try:
                _require_nonnegative_decimal(value, "monthly_values")
            except ValueError as error:
                raise ValueError(
                    "monthly_values must contain nonnegative finite Decimal values or None"
                ) from error


def _validate_period_inputs(
    quantity: Decimal | None, cost_won: int | None, label: str
) -> None:
    if (quantity is None) != (cost_won is None):
        raise ValueError(f"{label} quantity and cost must both be supplied or missing")
    if quantity is not None:
        _require_nonnegative_decimal(quantity, f"{label}_quantity")
        _require_nonnegative_int(
            cost_won,
            f"{label}_cost_won",
            maximum=MAX_SQLITE_INTEGER,
        )


def _validate_review_candidate(candidate: ReviewCandidate) -> None:
    if not isinstance(candidate, ReviewCandidate):
        raise ValueError("candidates must contain ReviewCandidate values")
    _require_nonnegative_int(candidate.destination_id, "destination_id", minimum=1)
    _require_nonnegative_int(candidate.display_order, "display_order")
    if not isinstance(candidate.name, str) or not candidate.name.strip():
        raise ValueError("candidate name must be nonblank text")
    if not isinstance(candidate.calculation, DestinationCalculation):
        raise ValueError("candidate calculation must be a DestinationCalculation")
    if candidate.calculation.kind is not CalculationKind.DESTINATION:
        raise ValueError("review candidates must use direct destination calculations")


def _require_unique_ids(values: Iterable[int], label: str) -> None:
    seen: set[int] = set()
    for value in values:
        if value in seen:
            raise ValueError(f"{label} must not contain duplicate destination ids")
        seen.add(value)


def _require_finite_decimal(value: object, field: str) -> None:
    if not isinstance(value, Decimal) or not value.is_finite():
        raise ValueError(f"{field} must be a finite Decimal")


def _require_nonnegative_decimal(value: object, field: str) -> None:
    _require_finite_decimal(value, field)
    if value < 0:
        raise ValueError(f"{field} must be a nonnegative finite Decimal")


def _require_nonnegative_int(
    value: object,
    field: str,
    *,
    minimum: int = 0,
    maximum: int | None = None,
) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < minimum
        or (maximum is not None and value > maximum)
    ):
        qualifier = "positive" if minimum else "nonnegative"
        upper_bound = "" if maximum is None else f" no greater than {maximum}"
        raise ValueError(f"{field} must be a {qualifier} integer{upper_bound}")


def _require_optional_nonnegative_decimal(value: object, field: str) -> None:
    if value is not None:
        _require_nonnegative_decimal(value, field)


def _require_optional_finite_decimal(value: object, field: str) -> None:
    if value is not None:
        _require_finite_decimal(value, field)


def _require_optional_nonnegative_int(value: object, field: str) -> None:
    if value is not None:
        _require_nonnegative_int(value, field)


def _require_optional_int(value: object, field: str) -> None:
    if value is not None and (isinstance(value, bool) or not isinstance(value, int)):
        raise ValueError(f"{field} must be an integer or None")
