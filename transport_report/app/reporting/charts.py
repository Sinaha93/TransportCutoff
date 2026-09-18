"""Deterministic chart data and PNG rendering for the August 2026 report."""

from __future__ import annotations

import json
import os
import tempfile
from collections.abc import Mapping
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from functools import lru_cache
from pathlib import Path
from typing import Literal

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.axes import Axes
from matplotlib.container import BarContainer
from matplotlib.figure import Figure
from matplotlib.ft2font import FT2Font
from matplotlib.font_manager import FontProperties
from matplotlib.lines import Line2D
from matplotlib.transforms import Bbox, IdentityTransform
from matplotlib.ticker import FuncFormatter, MultipleLocator
from PIL import Image, ImageChops

from app.domain.calculations import (
    HistoricalAverages,
    HistoricalValueKind,
    DestinationCalculation,
    GrandTotalResult,
    historical_averages,
)


REPORT_MONTH = "2026-08"
AVERAGE_LABELS = ("3개월 평균", "6개월 평균", "12개월 평균")
QUANTITY_CANVAS_PIXELS = (2125, 1441)
COST_CANVAS_PIXELS = (2060, 1145)
COMBINED_CANVAS_PIXELS = (2091, 933)


@dataclass(frozen=True, slots=True)
class PresentationAxis:
    maximum: Decimal
    major_interval: Decimal
    unit: str


QUANTITY_AXIS = PresentationAxis(Decimal("120000"), Decimal("20000"), "EA")
COST_AXIS = PresentationAxis(Decimal("60000"), Decimal("10000"), "천원")
COMBINED_COST_AXIS = PresentationAxis(Decimal("50000"), Decimal("5000"), "천원")

_QUANTITY_PLAN_BLUE = "#558ED5"
_COST_PLAN_BLUE = "#4F81BD"
_ACTUAL_GOLD = "#FFC000"
_ACTUAL_ORANGE = "#FF9900"
_COMPARISON_RED = "#C0504D"
_COMPARISON_GREEN = "#98B954"
_COMBINED_BLUE = "#0000FF"
_STANDARD_GRID = "#F2F2F2"
_STANDARD_BORDER = "#868686"
_COMBINED_BORDER = "#D9D9D9"
_COMBINED_PERIMETER = "#D8D8D8"
_KOREAN_FONT_FAMILY = "Malgun Gothic"
_REQUIRED_KOREAN_GLYPHS = "년월평균계획실적수량운반비천원"


@dataclass(frozen=True, slots=True)
class ChartReport:
    """Monthly totals required by all three charts.

    Plan mappings need the displayed August 2025 through August 2026 window.
    Actual mappings also need January through July 2025 because their full
    2025 values supply the prior-year comparison average. The combined chart
    consumes only the actual mappings.
    """

    report_month: str
    planned_quantity_by_month: Mapping[str, Decimal | None]
    actual_quantity_by_month: Mapping[str, Decimal | None]
    planned_cost_won_by_month: Mapping[str, int | Decimal | None]
    actual_cost_won_by_month: Mapping[str, int | Decimal | None]
    planned_quantity_history: HistoricalAverages | None = None
    actual_quantity_history: HistoricalAverages | None = None
    planned_cost_history: HistoricalAverages | None = None
    actual_cost_history: HistoricalAverages | None = None


@dataclass(frozen=True, slots=True)
class ChartSeries:
    label: str
    values: tuple[Decimal, ...]
    unit: str
    axis: Literal["left", "right"]
    presentation: Literal["bar", "line"]
    color: str
    line_style: str = "-"
    edge_color: str | None = None


@dataclass(frozen=True, slots=True)
class ChartData:
    chart_kind: Literal["quantity", "cost", "combined"]
    report_month: str
    labels: tuple[str, ...]
    comparison_label: str
    unit: str
    series: tuple[ChartSeries, ...]
    canvas_pixels: tuple[int, int]
    grid_color: str
    table_border_color: str
    axis_color: str
    text_color: str
    perimeter_color: str
    left_axis: PresentationAxis
    right_axis: PresentationAxis | None = None


def chart_report_from_calculation(
    *,
    report_month: str,
    current_total: DestinationCalculation,
    planned_quantity_by_month: Mapping[str, Decimal | None],
    actual_quantity_by_month: Mapping[str, Decimal | None],
    planned_cost_won_by_month: Mapping[str, int | Decimal | None],
    actual_cost_won_by_month: Mapping[str, int | Decimal | None],
    planned_quantity_history: HistoricalAverages,
    actual_quantity_history: HistoricalAverages,
    planned_cost_history: HistoricalAverages,
    actual_cost_history: HistoricalAverages,
) -> ChartReport:
    """Adapt a Task6 grand-total result and its precomputed histories.

    Historical mappings supply prior months; the canonical calculation result
    supplies the report-month totals. Supplied averages are accepted only when
    they exactly match Task6's authoritative recomputation from those mappings.
    """
    if not isinstance(current_total, GrandTotalResult):
        raise TypeError("current_total must be a Task6 GrandTotalResult")
    current_values = {
        "planned_quantity_by_month": current_total.planned_quantity,
        "actual_quantity_by_month": current_total.actual_quantity,
        "planned_cost_won_by_month": current_total.planned_cost_won,
        "actual_cost_won_by_month": current_total.actual_cost_won,
    }
    missing = [name for name, value in current_values.items() if value is None]
    if missing:
        raise ValueError(
            "current_total is incomplete for chart rendering: " + ", ".join(missing)
        )
    plan_quantity = _with_current(
        planned_quantity_by_month,
        report_month,
        current_total.planned_quantity,
        "planned_quantity_by_month",
    )
    actual_quantity = _with_current(
        actual_quantity_by_month,
        report_month,
        current_total.actual_quantity,
        "actual_quantity_by_month",
    )
    plan_cost = _with_current(
        planned_cost_won_by_month,
        report_month,
        current_total.planned_cost_won,
        "planned_cost_won_by_month",
    )
    actual_cost = _with_current(
        actual_cost_won_by_month,
        report_month,
        current_total.actual_cost_won,
        "actual_cost_won_by_month",
    )
    validated_histories = (
        _validated_history(
            report_month,
            plan_quantity,
            "planned_quantity_by_month",
            HistoricalValueKind.QUANTITY,
            planned_quantity_history,
        ),
        _validated_history(
            report_month,
            actual_quantity,
            "actual_quantity_by_month",
            HistoricalValueKind.QUANTITY,
            actual_quantity_history,
        ),
        _validated_history(
            report_month,
            plan_cost,
            "planned_cost_won_by_month",
            HistoricalValueKind.MONEY,
            planned_cost_history,
        ),
        _validated_history(
            report_month,
            actual_cost,
            "actual_cost_won_by_month",
            HistoricalValueKind.MONEY,
            actual_cost_history,
        ),
    )
    return ChartReport(
        report_month=report_month,
        planned_quantity_by_month=plan_quantity,
        actual_quantity_by_month=actual_quantity,
        planned_cost_won_by_month=plan_cost,
        actual_cost_won_by_month=actual_cost,
        planned_quantity_history=validated_histories[0],
        actual_quantity_history=validated_histories[1],
        planned_cost_history=validated_histories[2],
        actual_cost_history=validated_histories[3],
    )


def build_quantity_chart_data(report: ChartReport) -> ChartData:
    """Build quantity chart values without performing rendering work."""
    _require_supported_report(report)
    plan_history = _complete_history(
        report.report_month,
        report.planned_quantity_by_month,
        "planned_quantity_by_month",
        HistoricalValueKind.QUANTITY,
        require_comparison=False,
        supplied=report.planned_quantity_history,
    )
    actual_history = _complete_history(
        report.report_month,
        report.actual_quantity_by_month,
        "actual_quantity_by_month",
        HistoricalValueKind.QUANTITY,
        supplied=report.actual_quantity_history,
    )
    months = _display_months(report.report_month)
    labels = _display_labels(months)
    comparison_label = _comparison_label(report.report_month)
    plan_values = _metric_values(
        report.planned_quantity_by_month, months, plan_history
    )
    actual_values = _metric_values(
        report.actual_quantity_by_month, months, actual_history
    )
    comparison_value = _required_average(
        actual_history.comparison_year, "actual_quantity_by_month"
    )
    data = ChartData(
        chart_kind="quantity",
        report_month=report.report_month,
        labels=labels,
        comparison_label=comparison_label,
        unit="EA",
        series=(
            ChartSeries(
                "계획 수량",
                plan_values,
                "EA",
                "left",
                "bar",
                _QUANTITY_PLAN_BLUE,
                edge_color=_ACTUAL_GOLD,
            ),
            ChartSeries(
                "실적 수량",
                actual_values,
                "EA",
                "left",
                "bar",
                _ACTUAL_GOLD,
                edge_color=_ACTUAL_GOLD,
            ),
            ChartSeries(
                f"{comparison_label} 수량",
                (comparison_value,) * len(labels),
                "EA",
                "left",
                "line",
                _COMPARISON_RED,
                "--",
            ),
        ),
        canvas_pixels=QUANTITY_CANVAS_PIXELS,
        grid_color=_STANDARD_GRID,
        table_border_color=_STANDARD_BORDER,
        axis_color=_STANDARD_BORDER,
        text_color="#000000",
        perimeter_color=_STANDARD_BORDER,
        left_axis=QUANTITY_AXIS,
    )
    _validate_fixed_axes(data)
    return data


def build_cost_chart_data(report: ChartReport) -> ChartData:
    """Build transport-cost chart values in thousand won."""
    _require_supported_report(report)
    plan_history = _complete_history(
        report.report_month,
        report.planned_cost_won_by_month,
        "planned_cost_won_by_month",
        HistoricalValueKind.MONEY,
        require_comparison=False,
        supplied=report.planned_cost_history,
    )
    actual_history = _complete_history(
        report.report_month,
        report.actual_cost_won_by_month,
        "actual_cost_won_by_month",
        HistoricalValueKind.MONEY,
        supplied=report.actual_cost_history,
    )
    months = _display_months(report.report_month)
    labels = _display_labels(months)
    comparison_label = _comparison_label(report.report_month)
    plan_values = _scaled_values(
        _metric_values(report.planned_cost_won_by_month, months, plan_history)
    )
    actual_values = _scaled_values(
        _metric_values(report.actual_cost_won_by_month, months, actual_history)
    )
    comparison_value = _required_average(
        actual_history.comparison_year, "actual_cost_won_by_month"
    ) / Decimal(1_000)
    data = ChartData(
        chart_kind="cost",
        report_month=report.report_month,
        labels=labels,
        comparison_label=comparison_label,
        unit="천원",
        series=(
            ChartSeries(
                "계획 운반비",
                plan_values,
                "천원",
                "left",
                "bar",
                _COST_PLAN_BLUE,
                edge_color=_COST_PLAN_BLUE,
            ),
            ChartSeries(
                "실적 운반비",
                actual_values,
                "천원",
                "left",
                "bar",
                _ACTUAL_ORANGE,
                edge_color=_ACTUAL_ORANGE,
            ),
            ChartSeries(
                f"{comparison_label} 운반비",
                (comparison_value,) * len(labels),
                "천원",
                "left",
                "line",
                _COMPARISON_GREEN,
            ),
        ),
        canvas_pixels=COST_CANVAS_PIXELS,
        grid_color=_STANDARD_GRID,
        table_border_color=_STANDARD_BORDER,
        axis_color=_STANDARD_BORDER,
        text_color="#000000",
        perimeter_color=_STANDARD_BORDER,
        left_axis=COST_AXIS,
    )
    _validate_fixed_axes(data)
    return data


def build_combined_chart_data(report: ChartReport) -> ChartData:
    """Build actual quantity and actual cost series with explicit units/axes."""
    _require_supported_report(report)
    quantity_history = _complete_history(
        report.report_month,
        report.actual_quantity_by_month,
        "actual_quantity_by_month",
        HistoricalValueKind.QUANTITY,
        supplied=report.actual_quantity_history,
    )
    cost_history = _complete_history(
        report.report_month,
        report.actual_cost_won_by_month,
        "actual_cost_won_by_month",
        HistoricalValueKind.MONEY,
        supplied=report.actual_cost_history,
    )
    months = _display_months(report.report_month)
    labels = _display_labels(months)
    comparison_label = _comparison_label(report.report_month)
    quantity_values = _metric_values(
        report.actual_quantity_by_month, months, quantity_history
    )
    quantity_average = _required_average(
        quantity_history.comparison_year, "actual_quantity_by_month"
    )
    cost_values = _scaled_values(
        _metric_values(report.actual_cost_won_by_month, months, cost_history)
    )
    cost_average = _required_average(
        cost_history.comparison_year, "actual_cost_won_by_month"
    ) / Decimal(1_000)
    data = ChartData(
        chart_kind="combined",
        report_month=report.report_month,
        labels=labels,
        comparison_label=comparison_label,
        unit="EA / 천원",
        series=(
            ChartSeries(
                "실적 수량",
                quantity_values,
                "EA",
                "left",
                "bar",
                _ACTUAL_GOLD,
                edge_color=_ACTUAL_GOLD,
            ),
            ChartSeries(
                f"{comparison_label} 수량",
                (quantity_average,) * len(labels),
                "EA",
                "left",
                "line",
                _COMBINED_BLUE,
                "--",
            ),
            ChartSeries(
                "실적 운반비",
                cost_values,
                "천원",
                "right",
                "line",
                _COMPARISON_RED,
            ),
            ChartSeries(
                f"{comparison_label} 운반비",
                (cost_average,) * len(labels),
                "천원",
                "right",
                "line",
                "#92D050",
                "--",
            ),
        ),
        canvas_pixels=COMBINED_CANVAS_PIXELS,
        grid_color=_COMBINED_BORDER,
        table_border_color=_COMBINED_BORDER,
        axis_color="#595959",
        text_color="#595959",
        perimeter_color=_COMBINED_PERIMETER,
        left_axis=QUANTITY_AXIS,
        right_axis=COMBINED_COST_AXIS,
    )
    _validate_fixed_axes(data)
    return data


def render_quantity_chart(
    report: ChartReport, path: str | Path, dpi: int = 200
) -> None:
    _render_grouped_chart(build_quantity_chart_data(report), path, dpi)


def render_cost_chart(
    report: ChartReport, path: str | Path, dpi: int = 200
) -> None:
    _render_grouped_chart(build_cost_chart_data(report), path, dpi)


def render_combined_chart(
    report: ChartReport, path: str | Path, dpi: int = 200
) -> None:
    _render_combined_chart(build_combined_chart_data(report), path, dpi)


def _render_grouped_chart(data: ChartData, path: str | Path, dpi: int) -> None:
    font = _font(7.5)
    bold_font = _font(8, bold=True)
    figure, axis = _new_figure(data.canvas_pixels, dpi, data.perimeter_color)
    try:
        x_values = list(range(len(data.labels)))
        width = 0.28
        first, second, comparison = data.series
        first_bars = axis.bar(
            [value - width / 2 for value in x_values],
            _floats(first.values),
            width,
            color=first.color,
            edgecolor=first.edge_color,
            linewidth=0.8,
            label=first.label,
            zorder=3,
        )
        second_bars = axis.bar(
            [value + width / 2 for value in x_values],
            _floats(second.values),
            width,
            color=second.color,
            edgecolor=second.edge_color,
            linewidth=0.8,
            label=second.label,
            zorder=3,
        )
        comparison_line = axis.plot(
            x_values,
            _floats(comparison.values),
            color=comparison.color,
            linestyle=comparison.line_style,
            linewidth=2.4,
            dash_capstyle="round",
            label=comparison.label,
            zorder=4,
        )[0]
        _style_axis(axis, data, font)
        _set_axis_limit(axis, data.left_axis)
        axis.legend(
            [first_bars, second_bars, comparison_line],
            [first.label, second.label, comparison.label],
            loc="upper center",
            bbox_to_anchor=(0.5, 1.08),
            ncol=3,
            frameon=False,
            prop=bold_font,
        )
        figure.subplots_adjust(left=0.16, right=0.99, top=0.90, bottom=0.20)
        _add_table(axis, data, font, bold_font, y=-0.205, height=0.19)
        _label_grouped_bars(
            axis,
            (
                (first_bars, first.values, first.unit),
                (second_bars, second.values, second.unit),
            ),
            bold_font,
            data.text_color,
        )
        _save(figure, data, path, dpi)
    finally:
        plt.close(figure)


def _render_combined_chart(data: ChartData, path: str | Path, dpi: int) -> None:
    font = _font(7)
    bold_font = _font(7.5, bold=True)
    figure, left_axis = _new_figure(data.canvas_pixels, dpi, data.perimeter_color)
    try:
        right_axis = left_axis.twinx()
        x_values = list(range(len(data.labels)))
        quantity, quantity_average, cost, cost_average = data.series
        quantity_bars = left_axis.bar(
            x_values,
            _floats(quantity.values),
            width=0.40,
            color=quantity.color,
            edgecolor=quantity.edge_color,
            linewidth=0.8,
            label=quantity.label,
            zorder=3,
        )
        quantity_line = left_axis.plot(
            x_values,
            _floats(quantity_average.values),
            color=quantity_average.color,
            linestyle=quantity_average.line_style,
            linewidth=2.2,
            dash_capstyle="round",
            label=quantity_average.label,
            zorder=4,
        )[0]
        cost_line = right_axis.plot(
            x_values,
            _floats(cost.values),
            color=cost.color,
            linestyle=cost.line_style,
            linewidth=2.2,
            label=cost.label,
            zorder=5,
        )[0]
        cost_average_line = right_axis.plot(
            x_values,
            _floats(cost_average.values),
            color=cost_average.color,
            linestyle=cost_average.line_style,
            linewidth=2.2,
            dash_capstyle="round",
            label=cost_average.label,
            zorder=5,
        )[0]
        _style_axis(left_axis, data, font)
        right_axis.grid(False)
        right_axis.yaxis.set_major_formatter(FuncFormatter(_axis_number))
        right_axis.tick_params(
            axis="y", length=0, labelsize=8, colors=data.text_color
        )
        for label in right_axis.get_yticklabels():
            label.set_fontproperties(font)
        for spine in right_axis.spines.values():
            spine.set_visible(False)
        _set_axis_limit(left_axis, data.left_axis)
        if data.right_axis is None:
            raise ValueError("combined chart requires a fixed right axis")
        _set_axis_limit(right_axis, data.right_axis)
        left_axis.legend(
            [quantity_bars, quantity_line, cost_line, cost_average_line],
            [
                quantity.label,
                quantity_average.label,
                cost.label,
                cost_average.label,
            ],
            loc="upper center",
            bbox_to_anchor=(0.5, 1.10),
            ncol=4,
            frameon=False,
            prop=bold_font,
        )
        figure.subplots_adjust(left=0.16, right=0.93, top=0.88, bottom=0.25)
        _add_table(left_axis, data, font, bold_font, y=-0.275, height=0.255)
        _save(figure, data, path, dpi)
    finally:
        plt.close(figure)


def _new_figure(
    canvas_pixels: tuple[int, int], dpi: int, perimeter_color: str
) -> tuple[Figure, Axes]:
    if not isinstance(dpi, int) or isinstance(dpi, bool) or dpi <= 0:
        raise ValueError("dpi must be a positive integer")
    width, height = canvas_pixels
    figure, axis = plt.subplots(
        figsize=(width / dpi, height / dpi),
        dpi=dpi,
        facecolor="white",
    )
    figure.patch.set_edgecolor(perimeter_color)
    figure.patch.set_linewidth(1)
    return figure, axis


def _style_axis(axis: Axes, data: ChartData, font: FontProperties) -> None:
    axis.set_axisbelow(True)
    axis.grid(axis="y", color=data.grid_color, linewidth=0.7)
    axis.set_xticks(range(len(data.labels)))
    axis.set_xticklabels(())
    axis.tick_params(axis="x", length=0)
    axis.tick_params(axis="y", length=0, labelsize=9, colors=data.text_color)
    axis.yaxis.set_major_formatter(FuncFormatter(_axis_number))
    for label in axis.get_yticklabels():
        label.set_fontproperties(font)
    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.spines["left"].set_color(data.axis_color)
    axis.spines["bottom"].set_color(data.axis_color)


def _set_axis_limit(axis: Axes, presentation: PresentationAxis) -> None:
    axis.set_ylim(0, float(presentation.maximum))
    axis.yaxis.set_major_locator(MultipleLocator(float(presentation.major_interval)))


def _label_grouped_bars(
    axis: Axes,
    groups: tuple[
        tuple[BarContainer, tuple[Decimal, ...], str],
        tuple[BarContainer, tuple[Decimal, ...], str],
    ],
    font: FontProperties,
    text_color: str,
) -> None:
    figure = axis.figure
    figure.canvas.draw()
    renderer = figure.canvas.get_renderer()
    legend = axis.get_legend()
    blocked = [legend.get_window_extent(renderer)] if legend is not None else []
    if axis.tables:
        blocked.append(
            Bbox.union(
                [
                    cell.get_window_extent(renderer)
                    for cell in axis.tables[0].get_celld().values()
                ]
            )
        )
    placed: list[Bbox] = []
    candidates = (
        (3, "bottom"),
        (14, "bottom"),
        (25, "bottom"),
        (36, "bottom"),
        (-3, "top"),
        (-14, "top"),
        (-25, "top"),
        (-36, "top"),
    )
    for bars, values, unit in groups:
        for bar, value in zip(bars, values, strict=True):
            label = axis.annotate(
                _format_value(value, unit),
                xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                xytext=(0, candidates[0][0]),
                textcoords="offset points",
                ha="center",
                va=candidates[0][1],
                fontproperties=font,
                color=text_color,
                clip_on=False,
                zorder=6,
            )
            label.set_gid("bar-data-label")
            for offset, vertical_alignment in candidates:
                label.set_position((0, offset))
                label.set_verticalalignment(vertical_alignment)
                box = label.get_window_extent(renderer)
                inside_canvas = (
                    box.x0 >= figure.bbox.x0
                    and box.x1 <= figure.bbox.x1
                    and box.y0 >= figure.bbox.y0
                    and box.y1 <= figure.bbox.y1
                )
                if inside_canvas and not any(
                    box.overlaps(other) for other in (*blocked, *placed)
                ):
                    placed.append(box)
                    break
            else:
                raise RuntimeError(
                    f"no collision-free data-label position for {_format_value(value, unit)}"
                )
    figure.canvas.draw()


def _add_table(
    axis: Axes,
    data: ChartData,
    font: FontProperties,
    bold_font: FontProperties,
    *,
    y: float,
    height: float,
) -> None:
    # The source deck uses a compact header independently of its larger,
    # bold row labels.  These fixed sizes keep all 16 unabbreviated labels
    # inside their cells at the respective presentation canvas widths.
    header_font = _font(6.5 if data.chart_kind == "combined" else 7, bold=True)
    table = axis.table(
        cellText=[
            [_format_value(value, series.unit) for value in series.values]
            for series in data.series
        ],
        rowLabels=[series.label for series in data.series],
        colLabels=data.labels,
        cellLoc="center",
        rowLoc="left",
        bbox=(0, y, 1, height),
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font.get_size_in_points())
    for (row, column), cell in table.get_celld().items():
        cell.set_facecolor("white")
        cell.set_edgecolor(data.table_border_color)
        cell.set_linewidth(0.6)
        cell.PAD = 0.25 if column == -1 else 0.02
        cell.get_text().set_fontproperties(
            header_font if row == 0 else bold_font if column == -1 else font
        )
        cell.get_text().set_color("#000000" if column == -1 else data.text_color)
    axis.figure.canvas.draw()
    _add_table_keys(axis, data, table)
    axis.figure.canvas.draw()


def _add_table_keys(axis: Axes, data: ChartData, table) -> None:
    renderer = axis.figure.canvas.get_renderer()
    for index, series in enumerate(data.series, start=1):
        cell = table.get_celld()[(index, -1)]
        cell_box = cell.get_window_extent(renderer)
        text_box = cell.get_text().get_window_extent(renderer)
        left = cell_box.x0 + 4
        right = text_box.x0 - 4
        if right <= left:
            raise RuntimeError(f"table key space is unavailable for {series.label}")
        center_y = (cell_box.y0 + cell_box.y1) / 2
        if series.presentation == "bar":
            marker_points = min(5.0, (right - left) * 72 / axis.figure.dpi)
            key = Line2D(
                [(left + right) / 2],
                [center_y],
                marker="s",
                markersize=marker_points,
                linestyle="none",
                color=series.color,
                transform=IdentityTransform(),
                clip_on=False,
            )
        else:
            key = Line2D(
                [left, right],
                [center_y, center_y],
                color=series.color,
                linestyle=series.line_style,
                linewidth=2,
                transform=IdentityTransform(),
                clip_on=False,
            )
        key.set_gid("table-series-key")
        axis.add_artist(key)


def _save(figure: Figure, data: ChartData, path: str | Path, dpi: int) -> None:
    metadata = {
        "chart_kind": data.chart_kind,
        "report_month": data.report_month,
        "unit": data.unit,
        "legend_labels": [series.label for series in data.series],
        "labels": list(data.labels),
        "comparison_label": data.comparison_label,
        "final_month_label": data.labels[12],
        "final_month_values": {
            series.label: str(series.values[12])
            for series in data.series
            if not series.label.startswith(data.comparison_label)
        },
    }
    destination = Path(path)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.",
        suffix=".tmp",
        dir=destination.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        figure.savefig(
            temporary,
            format="png",
            dpi=dpi,
            facecolor="white",
            edgecolor=data.perimeter_color,
            metadata={
                "Title": (
                    f"{data.report_month} {data.chart_kind} transport report chart"
                ),
                "Description": json.dumps(
                    metadata,
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                ),
            },
        )
        _validate_png(temporary, data.canvas_pixels)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def _validate_png(path: Path, expected_size: tuple[int, int]) -> None:
    with Image.open(path) as image:
        image.verify()
    with Image.open(path) as image:
        if image.format != "PNG":
            raise ValueError(f"rendered chart is not PNG: {path}")
        if image.size != expected_size:
            raise ValueError(
                f"rendered chart size {image.size} does not match {expected_size}"
            )
        rgb = image.convert("RGB")
        white = Image.new("RGB", image.size, "white")
        if ImageChops.difference(rgb, white).getbbox() is None:
            raise ValueError("rendered chart is blank")


def _complete_history(
    report_month: str,
    values: Mapping[str, Decimal | int | None],
    field_name: str,
    value_kind: HistoricalValueKind,
    *,
    require_comparison: bool = True,
    supplied: HistoricalAverages | None = None,
) -> HistoricalAverages:
    averages = _validated_history(
        report_month,
        values,
        field_name,
        value_kind,
        supplied,
    )
    required_results = [
        averages.three_month,
        averages.six_month,
        averages.twelve_month,
    ]
    if require_comparison:
        required_results.append(averages.comparison_year)
    missing = {
        month
        for result in required_results
        for month in result.missing_months
    }
    if values.get(report_month) is None:
        missing.add(report_month)
    if missing:
        raise ValueError(
            f"{field_name} is incomplete for chart rendering; missing months: "
            + ", ".join(sorted(missing))
        )
    return averages


def _validated_history(
    report_month: str,
    values: Mapping[str, Decimal | int | None],
    field_name: str,
    value_kind: HistoricalValueKind,
    supplied: HistoricalAverages | None,
) -> HistoricalAverages:
    if supplied is not None and not isinstance(supplied, HistoricalAverages):
        raise TypeError(f"{field_name} history must be HistoricalAverages")
    try:
        authoritative = historical_averages(
            report_month,
            values,
            value_kind=value_kind,
        )
    except ValueError as error:
        raise ValueError(f"{field_name} is invalid: {error}") from error
    if supplied is not None and supplied != authoritative:
        raise ValueError(
            f"{field_name} history does not match {report_month} "
            f"{value_kind.value} monthly values"
        )
    return authoritative


def _with_current(
    values: Mapping[str, Decimal | int | None],
    report_month: str,
    current: Decimal | int | None,
    field_name: str,
) -> dict[str, Decimal | int | None]:
    if current is None:
        raise ValueError(f"current_total {field_name} is missing")
    result = dict(values)
    existing = result.get(report_month)
    if existing is not None and existing != current:
        raise ValueError(
            f"{field_name} {report_month} conflicts with current_total"
        )
    result[report_month] = current
    return result


def _validate_fixed_axes(data: ChartData) -> None:
    for series in data.series:
        presentation = data.left_axis if series.axis == "left" else data.right_axis
        if presentation is None:
            raise ValueError(f"{data.chart_kind} chart is missing its {series.axis} axis")
        maximum = max(series.values, default=Decimal(0))
        if maximum > presentation.maximum:
            raise ValueError(
                f"{data.chart_kind} chart {series.label} value {maximum} exceeds "
                f"fixed {presentation.unit} axis maximum "
                f"{presentation.maximum:,.0f}"
            )


def _metric_values(
    values: Mapping[str, Decimal | int | None],
    months: tuple[str, ...],
    averages: HistoricalAverages,
) -> tuple[Decimal, ...]:
    monthly = tuple(_required_month(values, month) for month in months)
    trailing = tuple(
        _required_average(result, "chart history")
        for result in (
            averages.three_month,
            averages.six_month,
            averages.twelve_month,
        )
    )
    return (*monthly, *trailing)


def _required_month(
    values: Mapping[str, Decimal | int | None], month: str
) -> Decimal:
    value = values.get(month)
    if value is None:
        raise ValueError(f"chart history is missing {month}")
    return value if isinstance(value, Decimal) else Decimal(value)


def _required_average(result, field_name: str) -> Decimal:
    if not result.complete or result.value is None:
        raise ValueError(f"{field_name} has incomplete history")
    return result.value


def _scaled_values(values: tuple[Decimal, ...]) -> tuple[Decimal, ...]:
    return tuple(value / Decimal(1_000) for value in values)


def _require_supported_report(report: ChartReport) -> None:
    if not isinstance(report, ChartReport):
        raise TypeError("report must be a ChartReport")
    if report.report_month != REPORT_MONTH:
        raise ValueError(
            f"chart rendering currently supports only {REPORT_MONTH}; "
            f"received {report.report_month!r}"
        )


def _display_months(report_month: str) -> tuple[str, ...]:
    year, month = (int(part) for part in report_month.split("-"))
    months = []
    for offset in range(12, -1, -1):
        zero_based = year * 12 + month - 1 - offset
        month_year, month_index = divmod(zero_based, 12)
        months.append(f"{month_year:04d}-{month_index + 1:02d}")
    return tuple(months)


def _display_labels(months: tuple[str, ...]) -> tuple[str, ...]:
    return (
        *(_month_label(month) for month in months),
        *AVERAGE_LABELS,
    )


def _month_label(month: str) -> str:
    year, number = month.split("-")
    return f"{int(year) % 100:02d}년 {int(number)}월"


def _comparison_label(report_month: str) -> str:
    return f"{int(report_month[:4]) % 100 - 1:02d}년 평균"


def _floats(values: tuple[Decimal, ...]) -> list[float]:
    return [float(value) for value in values]


def _axis_number(value: float, _position: int) -> str:
    return f"{value:,.0f}"


def _format_value(value: Decimal, unit: str) -> str:
    rounded = value.quantize(Decimal(1), rounding=ROUND_HALF_UP)
    return f"{int(rounded):,}"


@lru_cache(maxsize=1)
def _font_path() -> str:
    try:
        path = font_manager.findfont(
            FontProperties(family=[_KOREAN_FONT_FAMILY]),
            fallback_to_default=False,
        )
        character_map = FT2Font(path).get_charmap()
    except (OSError, RuntimeError, ValueError) as error:
        raise RuntimeError(
            "Malgun Gothic with Korean glyphs is required to render charts. "
            "Install or enable the Windows Korean supplemental fonts and retry."
        ) from error
    missing = [
        glyph
        for glyph in _REQUIRED_KOREAN_GLYPHS
        if ord(glyph) not in character_map
    ]
    if missing:
        raise RuntimeError(
            "Malgun Gothic does not contain the required Korean glyphs "
            f"({''.join(missing)}). Install or repair the Windows Korean "
            "supplemental fonts and retry."
        )
    return path


def _font(size: float, *, bold: bool = False) -> FontProperties:
    return FontProperties(
        fname=_font_path(),
        size=size,
        weight="bold" if bold else "normal",
    )
