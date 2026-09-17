from dataclasses import replace
from decimal import Decimal, Inexact, localcontext

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
        1: calculate_destination(Decimal("10"), 1_000, Decimal("12"), 1_200, destination_id=1),
        2: calculate_destination(Decimal("20"), 2_000, Decimal("18"), 1_800, destination_id=2),
        3: calculate_destination(Decimal("30"), 3_000, Decimal("30"), 3_300, destination_id=3),
        4: calculate_destination(Decimal("40"), 4_000, Decimal("50"), 5_000, destination_id=4),
    }
    members = [
        GroupMember(10, item.id, item.name, item.display_order, True, True)
        for item in destinations[:3]
    ]

    group = calculate_group(calculations, members)
    total = calculate_total(calculations, destinations)

    assert group.planned_quantity == Decimal("60")
    assert group.actual_quantity == Decimal("60")
    assert group.planned_cost_won == 6_000
    assert group.actual_cost_won == 6_300
    assert total.planned_quantity == Decimal("60")
    assert total.actual_quantity == Decimal("60")
    assert total.planned_cost_won == 10_000
    assert total.actual_cost_won == 11_300
    with pytest.raises(ValueError, match="derived group"):
        calculate_total({**calculations, 1: group}, destinations)
    with pytest.raises(ValueError, match="direct destination"):
        calculate_total({**calculations, 10: group}, destinations)


def test_group_member_metric_flags_are_honored():
    from app.domain.calculations import (
        calculate_destination,
        calculate_group,
    )

    calculations = {
        1: calculate_destination(Decimal("10"), 1_000, Decimal("20"), 2_000, destination_id=1),
        2: calculate_destination(Decimal("30"), 3_000, Decimal("40"), 4_000, destination_id=2),
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
    direct = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=2
    )
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
        1: calculate_destination(Decimal("0"), 0, Decimal("0"), 0, destination_id=1),
        2: calculate_destination(None, None, None, None, destination_id=2),
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
    from app.domain.calculations import CALCULATION_PRECISION, historical_averages

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
    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        expected_twelve_month = Decimal("218") / Decimal("12")
    assert result.twelve_month.value == expected_twelve_month
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
            calculate_destination(
                Decimal("1"), 100, Decimal("1"), 114, destination_id=3
            ),
        ),
        ReviewCandidate(
            2,
            "Positive boundary",
            20,
            calculate_destination(
                Decimal("1"), 100, Decimal("1"), 115, destination_id=2
            ),
        ),
        ReviewCandidate(
            1,
            "Negative boundary",
            10,
            calculate_destination(
                Decimal("1"), 100, Decimal("1"), 85, destination_id=1
            ),
        ),
    ]

    result = select_unit_cost_reviews(candidates)

    assert [item.destination_id for item in result.automatic_items] == [1, 2]
    assert [item.variance_pct for item in result.automatic_items] == [
        Decimal("-0.15"),
        Decimal("0.15"),
    ]
    assert result.validation_items == ()


def test_unit_cost_review_uses_unrounded_workbook_formula_chain():
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    calculation = calculate_destination(
        Decimal("1"), 1, Decimal("27"), 23, destination_id=1
    )

    assert calculation.planned_unit_cost == Decimal("1.00")
    assert calculation.actual_unit_cost == Decimal("0.85")
    assert calculation.actual_unit_cost_variance > Decimal("-0.15")
    assert calculation.actual_unit_cost_variance_pct > Decimal("-0.15")
    result = select_unit_cost_reviews(
        [ReviewCandidate(1, "Raw-ratio regression", 1, calculation)]
    )
    assert result.automatic_items == ()


@pytest.mark.parametrize(
    (
        "planned_cost_won",
        "planned_quantity",
        "actual_cost_won",
        "actual_quantity",
        "expected_pct",
        "selected",
    ),
    [
        (1, Decimal("1"), 17, Decimal("20"), Decimal("-0.15"), True),
        (1, Decimal("1"), 850_001, Decimal("1000000"), Decimal("-0.149999"), False),
        (1, Decimal("1"), 849_999, Decimal("1000000"), Decimal("-0.150001"), True),
        (1, Decimal("1"), 23, Decimal("20"), Decimal("0.15"), True),
        (1, Decimal("1"), 1_149_999, Decimal("1000000"), Decimal("0.149999"), False),
        (1, Decimal("1"), 1_150_001, Decimal("1000000"), Decimal("0.150001"), True),
    ],
)
def test_review_thresholds_follow_unrounded_workbook_values(
    planned_cost_won,
    planned_quantity,
    actual_cost_won,
    actual_quantity,
    expected_pct,
    selected,
):
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    calculation = calculate_destination(
        planned_quantity,
        planned_cost_won,
        actual_quantity,
        actual_cost_won,
        destination_id=1,
    )
    result = select_unit_cost_reviews(
        [ReviewCandidate(1, "Boundary", 1, calculation)]
    )

    assert calculation.actual_unit_cost_variance_pct == expected_pct
    assert bool(result.automatic_items) is selected


@pytest.mark.parametrize(
    ("planned_cost_won", "planned_quantity", "actual_cost_won", "actual_quantity"),
    [
        (5, Decimal("7"), 17, Decimal("28")),
        (1, Decimal("23"), 1, Decimal("20")),
    ],
)
def test_review_thresholds_include_exact_ratios_with_repeating_unit_costs(
    planned_cost_won, planned_quantity, actual_cost_won, actual_quantity
):
    from app.domain.calculations import (
        CALCULATION_PRECISION,
        DEFAULT_REVIEW_THRESHOLD,
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    calculation = calculate_destination(
        planned_quantity,
        planned_cost_won,
        actual_quantity,
        actual_cost_won,
        destination_id=1,
    )
    result = select_unit_cost_reviews(
        [ReviewCandidate(1, "Repeating exact boundary", 1, calculation)]
    )

    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        stored_magnitude = abs(calculation.actual_unit_cost_variance_pct)
    assert stored_magnitude < DEFAULT_REVIEW_THRESHOLD
    assert len(result.automatic_items) == 1


def test_negative_review_boundary_ignores_ambient_decimal_context():
    from app.domain.calculations import (
        ReviewCandidate,
        calculate_destination,
        select_unit_cost_reviews,
    )

    quantity = Decimal("830839729.9876")
    calculation = calculate_destination(
        quantity,
        3_359_510_071_930_000_560,
        quantity,
        2_855_583_561_140_500_476,
        destination_id=1,
    )
    with localcontext() as context:
        context.prec = 6
        context.traps[Inexact] = True
        result = select_unit_cost_reviews(
            [ReviewCandidate(1, "Large negative boundary", 1, calculation)]
        )

    assert len(result.automatic_items) == 1


def test_unit_cost_variance_percentage_preserves_workbook_formula_chain():
    from app.domain.calculations import CALCULATION_PRECISION, calculate_destination

    planned_cost_won = 7_921_731_534
    planned_quantity = Decimal("3475.1218")
    actual_cost_won = 1_806_341_205
    actual_quantity = Decimal("6862.2132")
    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        planned_raw = Decimal(planned_cost_won) / planned_quantity
        actual_raw = Decimal(actual_cost_won) / actual_quantity
        expected = (actual_raw - planned_raw) / planned_raw

    calculation = calculate_destination(
        planned_quantity,
        planned_cost_won,
        actual_quantity,
        actual_cost_won,
    )

    assert calculation.actual_unit_cost_variance_pct == expected


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
                calculate_destination(
                    Decimal("46"),
                    254_000,
                    Decimal("0"),
                    0,
                    destination_id=7,
                ),
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

    direct = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 115, destination_id=1
    )
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


def test_quantity_domain_accepts_business_boundaries_and_rejects_tiny_scale():
    from app.domain.calculations import (
        MAX_QUANTITY_EA,
        MAX_SQLITE_INTEGER,
        calculate_destination,
        unit_cost,
    )

    boundary = calculate_destination(
        MAX_QUANTITY_EA, 1, Decimal("0.0001"), 1
    )
    assert boundary.planned_quantity == MAX_QUANTITY_EA
    assert boundary.actual_quantity == Decimal("0.0001")
    assert unit_cost(1, Decimal("0.0001")) == Decimal("10000.00")
    with localcontext() as context:
        context.prec = 6
        largest_unit_cost = unit_cost(MAX_SQLITE_INTEGER, Decimal("0.0001"))
    assert largest_unit_cost == Decimal("92233720368547758070000.00")

    with pytest.raises(ValueError, match="at most 4 decimal places"):
        calculate_destination(Decimal("1E-28"), 1, Decimal("1"), 1)
    with pytest.raises(ValueError, match="at most 4 decimal places"):
        unit_cost(1, Decimal("0.00001"))
    with pytest.raises(ValueError, match="no greater than"):
        calculate_destination(MAX_QUANTITY_EA + Decimal("0.0001"), 1, Decimal("1"), 1)


def test_quantity_aggregate_preserves_scale_and_rejects_domain_overflow():
    from app.domain.calculations import calculate_destination, calculate_total

    destinations = [_destination(1, "First", 1), _destination(2, "Second", 2)]
    at_limit = {
        1: calculate_destination(
            Decimal("500000000.0001"), 1, Decimal("500000000.0001"), 1,
            destination_id=1,
        ),
        2: calculate_destination(
            Decimal("499999999.9999"), 1, Decimal("499999999.9999"), 1,
            destination_id=2,
        ),
    }
    overflow = {
        **at_limit,
        2: calculate_destination(
            Decimal("500000000.0000"), 1, Decimal("500000000.0000"), 1,
            destination_id=2,
        ),
    }

    assert calculate_total(at_limit, destinations).planned_quantity == Decimal(
        "1000000000.0000"
    )
    with pytest.raises(ValueError, match="planned_quantity.*no greater than"):
        calculate_total(overflow, destinations)


def test_history_uses_local_precision_and_validates_quantity_domain():
    from app.domain.calculations import CALCULATION_PRECISION, historical_averages

    values = {
        f"2025-{month:02d}": Decimal("1") for month in range(1, 13)
    }
    values["2025-12"] = Decimal("2")
    with localcontext() as context:
        context.prec = 6
        result = historical_averages("2026-01", values)

    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        expected = Decimal("4") / Decimal("3")
    assert result.three_month.value == expected

    invalid_scale = {**values, "2025-12": Decimal("1E-28")}
    with pytest.raises(ValueError, match="monthly_values.*at most 4 decimal places"):
        historical_averages("2026-01", invalid_scale)
    invalid_maximum = {**values, "2025-12": Decimal("1000000000.0001")}
    with pytest.raises(ValueError, match="monthly_values.*no greater than"):
        historical_averages("2026-01", invalid_maximum)


def test_decimal_arithmetic_errors_are_translated_to_domain_errors():
    from app.domain.calculations import variance

    with pytest.raises(ValueError, match="outside the supported Decimal domain"):
        variance(Decimal("9E+999999"), Decimal("-9E+999999"))


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

    with pytest.raises(TypeError, match="public factory"):
        replace(valid, **changes)


@pytest.mark.parametrize(
    ("changes", "message"),
    [
        ({"planned_unit_cost": Decimal("999")}, "planned_unit_cost"),
        ({"planned_unit_cost": 10.0}, "planned_unit_cost"),
        ({"actual_unit_cost": Decimal("999")}, "actual_unit_cost"),
        ({"quantity_variance": Decimal("999")}, "quantity_variance"),
        ({"quantity_variance_pct": Decimal("999")}, "quantity_variance_pct"),
        ({"cost_variance_won": 999}, "cost_variance_won"),
        ({"cost_variance_pct": Decimal("999")}, "cost_variance_pct"),
        ({"actual_unit_cost_variance": Decimal("999")}, "actual_unit_cost_variance"),
        (
            {"actual_unit_cost_variance_pct": Decimal("999")},
            "actual_unit_cost_variance_pct",
        ),
    ],
)
def test_result_records_reject_forged_calculated_fields(changes, message):
    from app.domain.calculations import calculate_destination

    valid = calculate_destination(
        Decimal("10"), 100, Decimal("20"), 240, destination_id=1
    )

    with pytest.raises(TypeError, match="public factory"):
        replace(valid, **changes)


def test_result_records_reject_kind_relabeling_without_matching_provenance():
    from app.domain.calculations import CalculationKind, calculate_destination, calculate_group

    destination = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )
    group = calculate_group(
        {1: destination},
        [GroupMember(10, 1, "Member", 1, True, True)],
    )

    with pytest.raises(TypeError):
        replace(group, kind=CalculationKind.DESTINATION)
    with pytest.raises(TypeError):
        replace(destination, kind=CalculationKind.DERIVED_GROUP)
    with pytest.raises(TypeError):
        replace(destination, kind=CalculationKind.GRAND_TOTAL)


def test_result_integrity_rejects_kind_and_provenance_replacement_together():
    from app.domain.calculations import (
        CalculationKind,
        DestinationProvenance,
        DerivedGroupProvenance,
        calculate_destination,
        calculate_group,
    )

    first = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )
    second = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=2
    )
    group = calculate_group(
        {1: first, 2: second},
        [
            GroupMember(10, 1, "First", 1, True, True),
            GroupMember(10, 2, "Second", 2, True, True),
        ],
    )

    with pytest.raises(TypeError):
        replace(
            group,
            kind=CalculationKind.DESTINATION,
            provenance=DestinationProvenance(1),
        )
    with pytest.raises(TypeError):
        replace(
            first,
            kind=CalculationKind.DERIVED_GROUP,
            provenance=DerivedGroupProvenance(10, (1,)),
        )


def test_result_identity_rejects_integrity_transplant_and_direct_replacement():
    from app.domain.calculations import (
        CalculationKind,
        calculate_destination,
        calculate_group,
    )

    destination = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )
    group = calculate_group(
        {1: destination},
        [GroupMember(10, 1, "Member", 1, True, True)],
    )

    assert not hasattr(destination, "_integrity")
    assert not hasattr(group, "_integrity")
    with pytest.raises(TypeError):
        replace(
            group,
            kind=CalculationKind.DESTINATION,
            provenance=destination.provenance,
            _integrity=object(),
        )
    with pytest.raises(TypeError):
        replace(group, provenance=destination.provenance)
    with pytest.raises(TypeError):
        replace(group, _integrity=object())


def test_result_backing_identity_fields_are_not_replaceable():
    from app.domain.calculations import (
        calculate_destination,
        calculate_group,
        calculate_total,
    )

    destination = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )
    group = calculate_group(
        {1: destination},
        [GroupMember(10, 1, "Member", 1, True, True)],
    )
    total = calculate_total(
        {1: destination},
        [_destination(1, "Destination", 1)],
    )

    for result, changes in (
        (destination, {"destination_id": 2}),
        (destination, {"_destination_id": 2}),
        (group, {"group_id": 20}),
        (group, {"_group_id": 20}),
        (group, {"member_destination_ids": (2,)}),
        (group, {"_member_destination_ids": (2,)}),
        (total, {"destination_ids": (2,)}),
        (total, {"_destination_ids": (2,)}),
    ):
        with pytest.raises((TypeError, ValueError)):
            replace(result, **changes)


@pytest.mark.parametrize("invalid_key", [True, "1", 0, -1])
def test_aggregate_rejects_invalid_calculation_map_keys(invalid_key):
    from app.domain.calculations import calculate_destination, calculate_total

    calculation = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )

    with pytest.raises(ValueError, match="calculation key"):
        calculate_total(
            {invalid_key: calculation},
            [_destination(1, "Destination", 1)],
        )


def test_aggregate_requires_map_key_to_match_destination_provenance():
    from app.domain.calculations import calculate_destination, calculate_total

    calculation = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=2
    )

    with pytest.raises(ValueError, match="does not match"):
        calculate_total({1: calculation}, [_destination(1, "Destination", 1)])


@pytest.mark.parametrize(
    ("changes", "message"),
    [
        ({"id": True}, "destination id"),
        ({"id": "1"}, "destination id"),
        ({"id": 0}, "destination id"),
        ({"include_quantity_total": 1}, "include_quantity_total"),
        ({"include_cost_total": "yes"}, "include_cost_total"),
        ({"include_sales_total": 1}, "include_sales_total"),
    ],
)
def test_total_rejects_invalid_master_rules(changes, message):
    from app.domain.calculations import calculate_destination, calculate_total

    rule = replace(_destination(1, "Destination", 1), **changes)
    calculation = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )

    with pytest.raises(ValueError, match=message):
        calculate_total({1: calculation}, [rule])


def test_total_rejects_duplicate_destination_rules():
    from app.domain.calculations import calculate_destination, calculate_total

    rule = _destination(1, "Destination", 1)
    calculation = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )

    with pytest.raises(ValueError, match="duplicate"):
        calculate_total({1: calculation}, [rule, rule])


def test_group_rejects_mixed_owner_ids_invalid_flags_and_duplicate_members():
    from app.domain.calculations import calculate_destination, calculate_group

    calculations = {
        1: calculate_destination(
            Decimal("1"), 100, Decimal("1"), 100, destination_id=1
        ),
        2: calculate_destination(
            Decimal("1"), 100, Decimal("1"), 100, destination_id=2
        ),
    }
    first = GroupMember(10, 1, "First", 1, True, True)
    wrong_owner = GroupMember(11, 2, "Second", 2, True, True)
    invalid_flag = GroupMember(10, 2, "Second", 2, 1, True)

    with pytest.raises(ValueError, match="same group_id"):
        calculate_group(calculations, [first, wrong_owner])
    with pytest.raises(ValueError, match="include_quantity"):
        calculate_group(calculations, [first, invalid_flag])
    with pytest.raises(ValueError, match="duplicate"):
        calculate_group(calculations, [first, first])


def test_group_rejects_bool_and_nonpositive_member_ids():
    from app.domain.calculations import calculate_destination, calculate_group

    calculation = calculate_destination(
        Decimal("1"), 100, Decimal("1"), 100, destination_id=1
    )

    with pytest.raises(ValueError, match="group_id"):
        calculate_group(
            {1: calculation},
            [GroupMember(True, 1, "Member", 1, True, True)],
        )
    with pytest.raises(ValueError, match="destination_id"):
        calculate_group(
            {1: calculation},
            [GroupMember(1, 0, "Member", 1, True, True)],
        )


def test_aggregate_cost_enforces_sqlite_bound_before_result_creation():
    from app.domain.calculations import MAX_SQLITE_INTEGER, calculate_destination, calculate_total

    destinations = [_destination(1, "First", 1), _destination(2, "Second", 2)]
    at_limit = {
        1: calculate_destination(
            Decimal("1"),
            MAX_SQLITE_INTEGER - 1,
            Decimal("1"),
            MAX_SQLITE_INTEGER - 1,
            destination_id=1,
        ),
        2: calculate_destination(
            Decimal("1"), 1, Decimal("1"), 1, destination_id=2
        ),
    }
    overflow = {
        1: calculate_destination(
            Decimal("1"),
            MAX_SQLITE_INTEGER,
            Decimal("1"),
            MAX_SQLITE_INTEGER,
            destination_id=1,
        ),
        2: calculate_destination(
            Decimal("1"), 1, Decimal("1"), 1, destination_id=2
        ),
    }

    assert calculate_total(at_limit, destinations).planned_cost_won == MAX_SQLITE_INTEGER
    with pytest.raises(ValueError, match="planned_cost_won"):
        calculate_total(overflow, destinations)


def test_aggregate_validates_cost_bound_before_unit_cost_arithmetic():
    from app.domain.calculations import MAX_SQLITE_INTEGER, calculate_destination, calculate_total

    destinations = [
        _destination(1, "Quantity", 1, include_quantity=True, include_cost=False),
        _destination(2, "Cost one", 2, include_quantity=False, include_cost=True),
        _destination(3, "Cost two", 3, include_quantity=False, include_cost=True),
    ]
    calculations = {
        1: calculate_destination(
            Decimal("0.0001"),
            0,
            Decimal("0.0001"),
            0,
            destination_id=1,
        ),
        2: calculate_destination(
            Decimal("1"),
            MAX_SQLITE_INTEGER,
            Decimal("1"),
            MAX_SQLITE_INTEGER,
            destination_id=2,
        ),
        3: calculate_destination(
            Decimal("1"), 1, Decimal("1"), 1, destination_id=3
        ),
    }

    with pytest.raises(ValueError, match="planned_cost_won"):
        calculate_total(calculations, destinations)
