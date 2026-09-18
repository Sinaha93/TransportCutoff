from __future__ import annotations

import json
from decimal import Decimal

import pytest
from PIL import Image, ImageChops


def _month_values(start: int, *, multiplier: int = 1):
    months = [
        *(f"2025-{month:02d}" for month in range(1, 13)),
        *(f"2026-{month:02d}" for month in range(1, 9)),
    ]
    return {
        month: Decimal(start + index * 10) * multiplier
        for index, month in enumerate(months)
    }


@pytest.fixture
def report():
    from app.reporting.charts import ChartReport

    return ChartReport(
        report_month="2026-08",
        planned_quantity_by_month=_month_values(1_000),
        actual_quantity_by_month=_month_values(2_000),
        planned_cost_won_by_month={
            month: int(value * 1_000)
            for month, value in _month_values(3_000).items()
        },
        actual_cost_won_by_month={
            month: int(value * 1_000)
            for month, value in _month_values(4_000).items()
        },
    )


def test_chart_data_uses_august_window_averages_and_comparison_semantics(report):
    from app.reporting.charts import (
        build_combined_chart_data,
        build_cost_chart_data,
        build_quantity_chart_data,
    )

    expected_labels = (
        "25년 8월",
        "25년 9월",
        "25년 10월",
        "25년 11월",
        "25년 12월",
        "26년 1월",
        "26년 2월",
        "26년 3월",
        "26년 4월",
        "26년 5월",
        "26년 6월",
        "26년 7월",
        "26년 8월",
        "3개월 평균",
        "6개월 평균",
        "12개월 평균",
    )

    quantity = build_quantity_chart_data(report)
    assert quantity.labels == expected_labels
    assert quantity.comparison_label == "25년 평균"
    assert quantity.unit == "EA"
    assert tuple(series.label for series in quantity.series) == (
        "계획 수량",
        "실적 수량",
        "25년 평균 수량",
    )
    assert tuple(series.unit for series in quantity.series) == ("EA", "EA", "EA")
    assert quantity.series[0].values[12] == Decimal("1190")
    assert quantity.series[1].values[12] == Decimal("2190")
    assert quantity.series[0].values[-3:] == (
        Decimal("1170"),
        Decimal("1155"),
        Decimal("1125"),
    )
    assert set(quantity.series[2].values) == {Decimal("2055")}

    cost = build_cost_chart_data(report)
    assert cost.labels == expected_labels
    assert cost.comparison_label == "25년 평균"
    assert cost.unit == "천원"
    assert tuple(series.label for series in cost.series) == (
        "계획 운반비",
        "실적 운반비",
        "25년 평균 운반비",
    )
    assert cost.series[0].values[12] == Decimal("3190")
    assert cost.series[1].values[12] == Decimal("4190")
    assert set(cost.series[2].values) == {Decimal("4055")}

    combined = build_combined_chart_data(report)
    assert combined.labels == expected_labels
    assert combined.unit == "EA / 천원"
    assert tuple((series.label, series.unit, series.axis) for series in combined.series) == (
        ("실적 수량", "EA", "left"),
        ("25년 평균 수량", "EA", "left"),
        ("실적 운반비", "천원", "right"),
        ("25년 평균 운반비", "천원", "right"),
    )
    assert combined.series[0].values[12] == Decimal("2190")
    assert combined.series[2].values[12] == Decimal("4190")


@pytest.mark.parametrize(
    ("renderer_name", "chart_kind", "expected_size", "expected_legends"),
    [
        (
            "render_quantity_chart",
            "quantity",
            (2125, 1441),
            ["계획 수량", "실적 수량", "25년 평균 수량"],
        ),
        (
            "render_cost_chart",
            "cost",
            (2060, 1145),
            ["계획 운반비", "실적 운반비", "25년 평균 운반비"],
        ),
        (
            "render_combined_chart",
            "combined",
            (2091, 933),
            ["실적 수량", "25년 평균 수량", "실적 운반비", "25년 평균 운반비"],
        ),
    ],
)
def test_renderers_write_stable_nonblank_png_with_chart_metadata(
    tmp_path,
    report,
    renderer_name,
    chart_kind,
    expected_size,
    expected_legends,
):
    from app.reporting import charts

    destination = tmp_path / f"{chart_kind}.png"
    getattr(charts, renderer_name)(report, destination)

    with Image.open(destination) as image:
        assert image.size == expected_size
        assert image.format == "PNG"
        rgb = image.convert("RGB")
        white = Image.new("RGB", image.size, "white")
        assert ImageChops.difference(rgb, white).getbbox() is not None
        metadata = json.loads(image.info["Description"])

    assert metadata["chart_kind"] == chart_kind
    assert metadata["report_month"] == "2026-08"
    assert metadata["legend_labels"] == expected_legends
    assert metadata["final_month_label"] == "26년 8월"
    assert metadata["final_month_values"]


def test_incomplete_history_is_actionable_and_legitimate_zero_is_preserved(report):
    from app.reporting.charts import ChartReport, build_quantity_chart_data

    missing = dict(report.actual_quantity_by_month)
    missing.pop("2025-01")
    incomplete = ChartReport(
        report_month=report.report_month,
        planned_quantity_by_month=report.planned_quantity_by_month,
        actual_quantity_by_month=missing,
        planned_cost_won_by_month=report.planned_cost_won_by_month,
        actual_cost_won_by_month=report.actual_cost_won_by_month,
    )

    with pytest.raises(ValueError, match=r"actual_quantity_by_month.*2025-01"):
        build_quantity_chart_data(incomplete)

    zero_values = dict(report.actual_quantity_by_month)
    zero_values["2026-08"] = Decimal(0)
    zero_report = ChartReport(
        report_month=report.report_month,
        planned_quantity_by_month=report.planned_quantity_by_month,
        actual_quantity_by_month=zero_values,
        planned_cost_won_by_month=report.planned_cost_won_by_month,
        actual_cost_won_by_month=report.actual_cost_won_by_month,
    )

    assert build_quantity_chart_data(zero_report).series[1].values[12] == 0


def test_plan_history_requires_only_the_displayed_thirteen_month_window(report):
    from dataclasses import replace

    from app.reporting.charts import build_cost_chart_data, build_quantity_chart_data

    displayed_months = {
        *(f"2025-{month:02d}" for month in range(8, 13)),
        *(f"2026-{month:02d}" for month in range(1, 9)),
    }
    report = replace(
        report,
        planned_quantity_by_month={
            month: value
            for month, value in report.planned_quantity_by_month.items()
            if month in displayed_months
        },
        planned_cost_won_by_month={
            month: value
            for month, value in report.planned_cost_won_by_month.items()
            if month in displayed_months
        },
    )

    assert build_quantity_chart_data(report).series[0].values[12] == Decimal("1190")
    assert build_cost_chart_data(report).series[0].values[12] == Decimal("3190")


def test_combined_chart_does_not_require_unused_plan_history(report):
    from dataclasses import replace

    from app.reporting.charts import build_combined_chart_data

    data = build_combined_chart_data(
        replace(
            report,
            planned_quantity_by_month={},
            planned_cost_won_by_month={},
        )
    )

    assert data.series[0].values[12] == Decimal("2190")
    assert data.series[2].values[12] == Decimal("4190")


def test_task6_total_and_histories_adapt_into_chart_data():
    from app.domain.calculations import (
        HistoricalValueKind,
        calculate_destination,
        calculate_total,
        historical_averages,
    )
    from app.domain.models import Destination
    from app.reporting.charts import (
        build_cost_chart_data,
        build_quantity_chart_data,
        chart_report_from_calculation,
    )

    destination = Destination(1, "Synthetic", 1, True, True, None, True, True, False)
    current = calculate_destination(
        Decimal("1234"),
        3_456_000,
        Decimal("2345"),
        4_567_000,
        destination_id=1,
    )
    total = calculate_total({1: current}, [destination])
    plan_quantity_history_values = _month_values(1_000)
    actual_quantity_history_values = _month_values(2_000)
    plan_cost_history_values = {
        month: int(value * 1_000) for month, value in _month_values(3_000).items()
    }
    actual_cost_history_values = {
        month: int(value * 1_000) for month, value in _month_values(4_000).items()
    }
    for values in (
        plan_quantity_history_values,
        actual_quantity_history_values,
        plan_cost_history_values,
        actual_cost_history_values,
    ):
        values.pop("2026-08")
    displayed = {
        *(f"2025-{month:02d}" for month in range(8, 13)),
        *(f"2026-{month:02d}" for month in range(1, 8)),
    }
    plan_quantity = {month: value for month, value in plan_quantity_history_values.items() if month in displayed}
    actual_quantity = {month: value for month, value in actual_quantity_history_values.items() if month in displayed}
    plan_cost = {month: value for month, value in plan_cost_history_values.items() if month in displayed}
    actual_cost = {month: value for month, value in actual_cost_history_values.items() if month in displayed}

    chart_report = chart_report_from_calculation(
        report_month="2026-08",
        current_total=total,
        planned_quantity_by_month=plan_quantity,
        actual_quantity_by_month=actual_quantity,
        planned_cost_won_by_month=plan_cost,
        actual_cost_won_by_month=actual_cost,
        planned_quantity_history=historical_averages(
            "2026-08", plan_quantity_history_values
        ),
        actual_quantity_history=historical_averages(
            "2026-08", actual_quantity_history_values
        ),
        planned_cost_history=historical_averages(
            "2026-08",
            plan_cost_history_values,
            value_kind=HistoricalValueKind.MONEY,
        ),
        actual_cost_history=historical_averages(
            "2026-08",
            actual_cost_history_values,
            value_kind=HistoricalValueKind.MONEY,
        ),
    )

    quantity = build_quantity_chart_data(chart_report)
    cost = build_cost_chart_data(chart_report)
    assert total.destination_ids == (1,)
    assert quantity.series[0].values[12] == Decimal("1234")
    assert quantity.series[1].values[12] == Decimal("2345")
    assert quantity.series[1].values[-3:] == (
        Decimal("2170"),
        Decimal("2155"),
        Decimal("2125"),
    )
    assert cost.series[0].values[12] == Decimal("3456")
    assert cost.series[1].values[12] == Decimal("4567")


def test_scope_is_limited_to_august_2026(report):
    from dataclasses import replace

    from app.reporting.charts import build_cost_chart_data

    with pytest.raises(ValueError, match="2026-08"):
        build_cost_chart_data(replace(report, report_month="2026-09"))


def test_values_above_fixed_reference_axis_fail_clearly(report):
    from dataclasses import replace

    from app.reporting.charts import build_quantity_chart_data

    values = dict(report.actual_quantity_by_month)
    values["2026-08"] = Decimal("120001")

    with pytest.raises(ValueError, match=r"quantity.*120,000"):
        build_quantity_chart_data(replace(report, actual_quantity_by_month=values))


def test_renderer_closes_figure_when_saving_fails(monkeypatch, tmp_path, report):
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure

    from app.reporting.charts import render_quantity_chart

    def fail_save(self, *args, **kwargs):
        raise OSError("forced save failure")

    monkeypatch.setattr(Figure, "savefig", fail_save)

    with pytest.raises(OSError, match="forced save failure"):
        render_quantity_chart(report, tmp_path / "quantity.png")

    assert plt.get_fignums() == []


@pytest.mark.parametrize(
    (
        "renderer_name",
        "expected_legends",
        "bar_face",
        "bar_edge",
        "line_colors",
        "grid_color",
        "table_border",
        "left_axis",
        "right_axis",
    ),
    [
        (
            "render_quantity_chart",
            ["계획 수량", "실적 수량", "25년 평균 수량"],
            "#558ed5",
            "#ffc000",
            ["#c0504d"],
            "#f2f2f2",
            "#868686",
            (120_000, 20_000),
            None,
        ),
        (
            "render_cost_chart",
            ["계획 운반비", "실적 운반비", "25년 평균 운반비"],
            "#4f81bd",
            "#4f81bd",
            ["#98b954"],
            "#f2f2f2",
            "#868686",
            (60_000, 10_000),
            None,
        ),
        (
            "render_combined_chart",
            ["실적 수량", "25년 평균 수량", "실적 운반비", "25년 평균 운반비"],
            "#ffc000",
            "#ffc000",
            ["#0000ff", "#c0504d", "#92d050"],
            "#d9d9d9",
            "#d9d9d9",
            (120_000, 20_000),
            (50_000, 5_000),
        ),
    ],
)
def test_rendered_artists_match_reference_palette_legend_and_table_keys(
    monkeypatch,
    tmp_path,
    report,
    renderer_name,
    expected_legends,
    bar_face,
    bar_edge,
    line_colors,
    grid_color,
    table_border,
    left_axis,
    right_axis,
):
    from matplotlib.colors import to_hex
    from matplotlib.figure import Figure

    from app.reporting import charts

    captured = {}

    def capture_figure(self, *args, **kwargs):
        captured["figure"] = self
        captured["renderer"] = self.canvas.get_renderer()

    monkeypatch.setattr(Figure, "savefig", capture_figure)

    getattr(charts, renderer_name)(report, tmp_path / "captured.png")

    figure = captured["figure"]
    axis = figure.axes[0]
    first_bar = axis.containers[0].patches[0]
    assert to_hex(first_bar.get_facecolor()) == bar_face
    assert to_hex(first_bar.get_edgecolor()) == bar_edge
    assert [
        to_hex(line.get_color())
        for axes in figure.axes
        for line in axes.lines
        if line.get_gid() != "table-series-key"
    ] == line_colors
    assert [text.get_text() for text in axis.get_legend().get_texts()] == expected_legends
    assert to_hex(axis.get_ygridlines()[0].get_color()) == grid_color

    table = axis.tables[0]
    assert to_hex(table.get_celld()[(1, 0)].get_edgecolor()) == table_border
    assert [
        to_hex(table.get_celld()[(row, -1)].get_text().get_color())
        for row in range(1, len(expected_legends) + 1)
    ] == ["#000000"] * len(expected_legends)
    table_keys = [
        artist
        for artist in axis.get_children()
        if artist.get_gid() == "table-series-key"
    ]
    assert [to_hex(key.get_color()) for key in table_keys] == [
        to_hex(series.color) for series in (
            charts.build_combined_chart_data(report).series
            if renderer_name == "render_combined_chart"
            else charts.build_cost_chart_data(report).series
            if renderer_name == "render_cost_chart"
            else charts.build_quantity_chart_data(report).series
        )
    ]
    assert all(not key.get_clip_on() for key in table_keys)
    assert figure.subplotpars.left >= 0.16
    renderer = captured["renderer"]
    figure_box = figure.bbox
    for row, key in enumerate(table_keys, start=1):
        cell_box = table.get_celld()[(row, -1)].get_window_extent(renderer)
        text_box = table.get_celld()[(row, -1)].get_text().get_window_extent(renderer)
        key_box = key.get_window_extent(renderer)
        assert figure_box.contains(cell_box.x0, cell_box.y0)
        assert figure_box.contains(cell_box.x1, cell_box.y1)
        assert cell_box.contains(key_box.x0, key_box.y0)
        assert cell_box.contains(key_box.x1, key_box.y1)
        assert key_box.x1 < text_box.x0
    assert axis.get_ylim()[1] == left_axis[0]
    assert axis.get_yticks()[1] - axis.get_yticks()[0] == left_axis[1]
    if right_axis is not None:
        assert figure.axes[1].get_ylim()[1] == right_axis[0]
        assert figure.axes[1].get_yticks()[1] - figure.axes[1].get_yticks()[0] == right_axis[1]
    else:
        labels = {text.get_text() for text in axis.texts}
        assert {"1,190", "2,190"} <= labels or {"3,190", "4,190"} <= labels
