from dataclasses import replace
from decimal import Decimal

import pytest

from app.domain.models import Destination, GroupMember


def test_unit_transport_cost_is_cost_divided_by_quantity():
    from app.domain.calculations import calculate_destination

    result = calculate_destination(
        planned_quantity=Decimal("15152"),
        planned_cost_won=1_578_333,
        actual_quantity=Decimal("14782"),
        actual_cost_won=1_766_000,
    )

    assert result.planned_unit_cost == Decimal("104.17")
    assert result.actual_unit_cost == Decimal("119.47")


def test_zero_quantity_produces_no_unit_cost_instead_of_division_error():
    from app.domain.calculations import calculate_destination

    result = calculate_destination(
        Decimal("46"), 254_000, Decimal("0"), 0
    )

    assert result.actual_unit_cost is None
    assert result.actual_unit_cost_variance_pct is None


def _destination(
    destination_id: int,
    name: str,
    display_order: int,
    *,
    include_quantity: bool = True,
    include_cost: bool = True,
) -> Destination:
    return Destination(
        id=destination_id,
        name=name,
        display_order=display_order,
        active=True,
        required_for_report=True,
        representative_item=None,
        include_quantity_total=include_quantity,
        include_cost_total=include_cost,
        include_sales_total=True,
    )


def test_totals_follow_master_rules_and_do_not_count_derived_group_twice():
    from app.domain.calculations import (
        calculate_destination,
        calculate_group,
        calculate_total,
    )

    destinations = [
        _destination(1, "현대 울산", 1),
        _destination(2, "포레시아 영천", 2),
        _destination(3, "세종공업", 3),
        _destination(4, "당진", 4, include_quantity=False, include_cost=True),
    ]
    calculations = {
        1: calculate_destination(Decimal("10"), 1_000, Decimal("12"), 1_200),
        2: calculate_destination(Decimal("20"), 2_000, Decimal("18"), 1_800),
        3: calculate_destination(Decimal("30"), 3_000, Decimal("30"), 3_300),
        4: calculate_destination(Decimal("40"), 4_000, Decimal("50"), 5_000),
    }
    members = [
        GroupMember(10, item.id, item.name, item.display_order, True, True)
        for item in destinations[:3]
    ]

    group = calculate_group(calculations, members)
    total = calculate_total(calculations, destinations)
    total_with_unrelated_derived_result = calculate_total(
        {**calculations, 10: group}, destinations
    )

    assert group.planned_quantity == Decimal("60")
    assert group.actual_quantity == Decimal("60")
    assert group.planned_cost_won == 6_000
    assert group.actual_cost_won == 6_300
    assert total.planned_quantity == Decimal("60")
    assert total.actual_quantity == Decimal("60")
    assert total.planned_cost_won == 10_000
    assert total.actual_cost_won == 11_300
    assert total_with_unrelated_derived_result == total

    with pytest.raises(ValueError, match="derived group"):
        calculate_total({**calculations, 1: group}, destinations)


def test_group_member_metric_flags_are_honored():
    from app.domain.calculations import (
        calculate_destination,
        calculate_group,
    )

    calculations = {
        1: calculate_destination(Decimal("10"), 1_000, Decimal("20"), 2_000),
        2: calculate_destination(Decimal("30"), 3_000, Decimal("40"), 4_000),
    }
    members = [
        GroupMember(1, 1, "First", 1, True, False),
        GroupMember(1, 2, "Second", 2, False, True),
    ]

    group = calculate_group(calculations, members)

    assert group.planned_quantity == Decimal("10")
    assert group.actual_quantity == Decimal("20")
    assert group.planned_cost_won == 3_000
    assert group.actual_cost_won == 4_000


def test_aggregate_validates_all_sources_before_missing_values_short_circuit():
    from app.domain.calculations import (
        calculate_destination,
        calculate_group,
        calculate_total,
    )

    destinations = [_destination(1, "Missing", 1), _destination(2, "Collides", 2)]
    direct = calculate_destination(Decimal("1"), 100, Decimal("1"), 100)
    derived = calculate_group(
        {2: direct},
        [GroupMember(2, 2, "Collides", 1, True, True)],
    )

    with pytest.raises(ValueError, match="derived group"):
        calculate_total({2: derived}, destinations)


def test_missing_included_metric_propagates_instead_of_becoming_zero():
    from app.domain.calculations import (
        calculate_destination,
        calculate_total,
    )

    calculations = {
        1: calculate_destination(Decimal("0"), 0, Decimal("0"), 0),
        2: calculate_destination(None, None, None, None),
    }
    total = calculate_total(
        calculations,
        [_destination(1, "Zero", 1), _destination(2, "Missing", 2)],
    )

    assert calculations[1].planned_quantity == Decimal("0")
    assert calculations[2].planned_quantity is None
    assert total.planned_quantity is None
    assert total.planned_cost_won is None


def test_history_windows_end_before_report_month_and_use_preceding_year():
    from app.domain.calculations import historical_averages

    values = {
        f"2025-{month:02d}": Decimal(month) for month in range(1, 13)
    }
    values.update(
        {f"2026-{month:02d}": Decimal(20 + month) for month in range(1, 8)}
    )

    result = historical_averages("2026-08", values)

    assert result.three_month.expected_months == (
        "2026-05",
        "2026-06",
        "2026-07",
    )
    assert result.six_month.expected_months == (
        "2026-02",
        "2026-03",
        "2026-04",
        "2026-05",
        "2026-06",
        "2026-07",
    )
    assert result.twelve_month.expected_months == (
        "2025-08",
        "2025-09",
        "2025-10",
        "2025-11",
        "2025-12",
        "2026-01",
        "2026-02",
        "2026-03",
        "2026-04",
        "2026-05",
        "2026-06",
        "2026-07",
    )
    assert result.comparison_year.expected_months == tuple(
        f"2025-{month:02d}" for month in range(1, 13)
    )
    assert result.three_month.value == Decimal("26")
    assert result.six_month.value == Decimal("24.5")
    assert result.twelve_month.value == Decimal("218") / Decimal("12")
    assert result.comparison_year.value == Decimal("6.5")


def test_history_window_rolls_across_year_and_reports_missing_months():
    from app.domain.calculations import historical_averages

    values = {
        f"2025-{month:02d}": Decimal("5") for month in range(1, 13)
    }
    del values["2025-11"]

    result = historical_averages("2026-01", values)

    assert result.three_month.expected_months == (
        "2025-10",
        "2025-11",
        "2025-12",
    )
    assert result.three_month.complete is False
    assert result.three_month.value is None
    assert result.three_month.missing_months == ("2025-11",)
    assert result.comparison_year.complete is False
    assert result.comparison_year.missing_months == ("2025-11",)


def test_history_treats_explicit_zero_as_present():
    from app.domain.calculations import historical_averages

    values = {
        f"2025-{month:02d}": Decimal("0") for month in range(1, 13)
    }

    result = historical_averages("2026-01", values)

    assert result.three_month.complete is True
    assert result.three_month.missing_months == ()
    assert result.three_month.value == Decimal("0")


def test_review_selector_includes_exact_thresholds_and_orders_deterministically():
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    candidates = [
        ReviewCandidate(
            3,
            "Below threshold",
            30,
            calculate_destination(Decimal("1"), 100, Decimal("1"), 114),
        ),
        ReviewCandidate(
            2,
            "Positive boundary",
            20,
            calculate_destination(Decimal("1"), 100, Decimal("1"), 115),
        ),
        ReviewCandidate(
            1,
            "Negative boundary",
            10,
            calculate_destination(Decimal("1"), 100, Decimal("1"), 85),
        ),
    ]

    result = select_unit_cost_reviews(candidates)

    assert [item.destination_id for item in result.automatic_items] == [1, 2]
    assert [item.variance_pct for item in result.automatic_items] == [
        Decimal("-0.15"),
        Decimal("0.15"),
    ]
    assert result.validation_items == ()


def test_review_selector_surfaces_unavailable_variance_without_narrative():
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    result = select_unit_cost_reviews(
        [
            ReviewCandidate(
                7,
                "Zero actual",
                3,
                calculate_destination(Decimal("46"), 254_000, Decimal("0"), 0),
            )
        ]
    )

    assert result.automatic_items == ()
    assert len(result.validation_items) == 1
    assert result.validation_items[0].destination_id == 7
    assert result.validation_items[0].code == "UNIT_COST_VARIANCE_UNAVAILABLE"


def test_review_selector_rejects_derived_group_candidates():
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        calculate_group,
        select_unit_cost_reviews,
    )

    direct = calculate_destination(Decimal("1"), 100, Decimal("1"), 115)
    derived = calculate_group(
        {1: direct},
        [GroupMember(1, 1, "Member", 1, True, True)],
    )

    with pytest.raises(ValueError, match="direct destination"):
        select_unit_cost_reviews([ReviewCandidate(1, "Group", 1, derived)])


def test_rounding_matches_excel_positive_half_away_from_zero():
    from app.domain.calculations import unit_cost

    assert unit_cost(201, Decimal("200")) == Decimal("1.01")


def test_destination_variances_use_actual_minus_plan():
    from app.domain.calculations import (
        calculate_destination,
        variance,
        variance_pct,
    )

    result = calculate_destination(Decimal("10"), 100, Decimal("15"), 180)

    assert result.quantity_variance == Decimal("5")
    assert result.quantity_variance_pct == Decimal("0.5")
    assert result.cost_variance_won == 80
    assert result.cost_variance_pct == Decimal("0.8")
    assert result.actual_unit_cost_variance == Decimal("2.00")
    assert result.actual_unit_cost_variance_pct == Decimal("0.2")
    assert variance(Decimal("3"), Decimal("5")) == Decimal("-2")
    assert variance_pct(Decimal("1"), Decimal("0")) is None


@pytest.mark.parametrize(
    ("args", "message"),
    [
        ((Decimal("-1"), 0, Decimal("0"), 0), "planned_quantity"),
        ((Decimal("1"), -1, Decimal("0"), 0), "planned_cost_won"),
        ((Decimal("NaN"), 0, Decimal("0"), 0), "planned_quantity"),
        ((1.0, 1, Decimal("0"), 0), "planned_quantity"),
        ((Decimal("1"), True, Decimal("0"), 0), "planned_cost_won"),
        ((Decimal("1"), 2**63, Decimal("0"), 0), "planned_cost_won"),
        ((None, 1, Decimal("0"), 0), "planned"),
    ],
)
def test_destination_rejects_invalid_or_inconsistent_inputs(args, message):
    from app.domain.calculations import calculate_destination

    with pytest.raises(ValueError, match=message):
        calculate_destination(*args)


@pytest.mark.parametrize(
    "values",
    [
        {"2025-12": Decimal("Infinity")},
        {"2025-12": 1.0},
        {"2025-12": Decimal("-1")},
    ],
)
def test_history_rejects_non_decimal_values(values):
    from app.domain.calculations import historical_averages

    with pytest.raises(ValueError, match="monthly_values"):
        historical_averages("2026-01", values)


def test_helpers_reject_bool_float_and_nonfinite_values():
    from app.domain.calculations import unit_cost, variance, variance_pct

    with pytest.raises(ValueError, match="cost_won"):
        unit_cost(True, Decimal("1"))
    with pytest.raises(ValueError, match="quantity_ea"):
        unit_cost(1, Decimal("Infinity"))
    with pytest.raises(ValueError, match="actual"):
        variance(1.0, Decimal("1"))
    with pytest.raises(ValueError, match="plan"):
        variance_pct(Decimal("1"), Decimal("NaN"))


@pytest.mark.parametrize(
    ("changes", "message"),
    [
        ({"planned_quantity": 1.0}, "planned_quantity"),
        ({"planned_cost_won": True}, "planned_cost_won"),
        ({"planned_cost_won": -1}, "planned_cost_won"),
        ({"planned_cost_won": None}, "planned"),
        ({"planned_cost_won": 2**63}, "planned_cost_won"),
        ({"actual_unit_cost": Decimal("Infinity")}, "actual_unit_cost"),
        ({"actual_unit_cost_variance_pct": 0.2}, "actual_unit_cost_variance_pct"),
    ],
)
def test_result_records_reject_invalid_numeric_fields(changes, message):
    from app.domain.calculations import calculate_destination

    valid = calculate_destination(Decimal("1"), 100, Decimal("1"), 100)

    with pytest.raises(ValueError, match=message):
        replace(valid, **changes)
