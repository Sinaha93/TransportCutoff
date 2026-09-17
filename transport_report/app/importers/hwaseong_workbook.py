from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook


REGULAR_SHEET = "화성운반비내역(8월)"
_REGULAR_SHEET = re.compile(r"^화성운반비내역\((?P<month>[1-9]|1[0-2])월\)$")
SUBCONTRACT_SHEETS = (
    "용차1-OK로지웰",
    "용차2-대원로지스틱",
    "용차3-정동물류",
)
_MONTH = re.compile(r"^[0-9]{4}-(0[1-9]|1[0-2])$")
_REGULAR_COLUMN_B_HEADERS = {"날짜", "운반지역", "운반지", "목적지", "납품처"}
_REGULAR_TOTAL_COUNT_HEADERS = {"계", "총회수", "총횟수"}
_REGULAR_UNIT_HEADERS = {"단가", "운반단가", "운임단가"}
_REGULAR_SUBTOTAL_HEADERS = {"소계", "금액", "운반비"}
_DATE_HEADERS = {"일자", "날짜", "오더일자", "운반일", "운송일"}
_DESTINATION_HEADERS = {"운반지역", "운반지", "목적지", "납품처", "도착지"}
_TONNAGE_HEADERS = {"톤수", "차량톤수", "차종", "차종(t)", "차종(톤)"}
_AMOUNT_HEADERS = {"금액", "운반비", "운임", "기본요금"}
_NUMBER_HEADERS = {"no", "번호"}
_TOTAL_LABELS = {"합계", "총계", "계"}
_SUBCONTRACT_TOTAL_LABEL_COLUMNS = (1, 7)
_SUBCONTRACT_LAST_FOOTPRINT_COLUMN = 10
_SQLITE_MAX = 2**63 - 1
# Conservative source-cell limits: 10,000 trips/day, 4 decimals, 16 canonical chars.
MAX_TRIP_COUNT = Decimal("10000")
MAX_TRIP_COUNT_SCALE = 4
MAX_TRIP_COUNT_CANONICAL_LENGTH = 16


class TransportWorkbookError(Exception):
    """Base error for workbooks that cannot be imported safely."""


class WorkbookStructureError(TransportWorkbookError):
    """Raised when a required workbook structure or value is missing."""


class SubtotalMismatchError(TransportWorkbookError):
    """Raised when a regular row does not reconcile to its cached subtotal."""


class TotalCountMismatchError(TransportWorkbookError):
    """Raised when a regular row does not reconcile to its cached trip count."""


class SubcontractTotalMismatchError(TransportWorkbookError):
    """Raised when subcontract details do not reconcile to the displayed total."""


@dataclass(frozen=True, slots=True)
class ParsedTransportEntry:
    report_month: str
    destination_alias: str
    source_sheet: str
    source_row: int
    source_date: date
    day: int
    transport_type: str
    trip_count: Decimal
    unit_rate_won: int | None
    cost_won: int
    vehicle_type: str | None = None
    vehicle_driver_group: str | None = None
    source_note: str | None = None


class HwaseongWorkbookParser:
    def parse(
        self, workbook_path: str | Path, report_month: str
    ) -> list[ParsedTransportEntry]:
        if not isinstance(report_month, str) or _MONTH.fullmatch(report_month) is None:
            raise WorkbookStructureError("report_month must use YYYY-MM")

        path = Path(workbook_path)
        workbook = None
        formula_workbook = None
        try:
            workbook = load_workbook(
                path,
                data_only=True,
                read_only=False,
                keep_vba=False,
                keep_links=False,
            )
            formula_workbook = load_workbook(
                path,
                data_only=False,
                read_only=False,
                keep_vba=False,
                keep_links=False,
            )
        except Exception as error:
            if workbook is not None:
                workbook.close()
            raise WorkbookStructureError(f"Workbook could not be opened: {error}") from error

        try:
            regular_sheets = [
                name for name in workbook.sheetnames if _REGULAR_SHEET.fullmatch(name)
            ]
            missing = [name for name in SUBCONTRACT_SHEETS if name not in workbook.sheetnames]
            if not regular_sheets:
                missing.insert(0, "화성운반비내역(<month>월)")
            if missing:
                raise WorkbookStructureError(
                    "Workbook is missing required sheet(s): " + ", ".join(missing)
                )
            if len(regular_sheets) != 1:
                raise WorkbookStructureError(
                    "Workbook must contain exactly one regular transport sheet"
                )

            regular_sheet = workbook[regular_sheets[0]]
            _validate_regular_sheet_month(regular_sheet.title, report_month)
            rows = self._parse_regular(
                regular_sheet,
                formula_workbook[regular_sheets[0]],
                report_month,
            )
            for sheet_name in SUBCONTRACT_SHEETS:
                rows.extend(
                    self._parse_subcontract(
                        workbook[sheet_name],
                        formula_workbook[sheet_name],
                        report_month,
                    )
                )
            return rows
        finally:
            workbook.close()
            formula_workbook.close()

    def _parse_regular(
        self, sheet, formula_sheet, report_month: str
    ) -> list[ParsedTransportEntry]:
        header_row = _find_regular_header(sheet)
        if header_row is None:
            raise WorkbookStructureError(
                f"Required header was not found in sheet {sheet.title}"
            )

        rows: list[ParsedTransportEntry] = []
        vehicle_driver_group: str | None = None
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if _is_regular_header(sheet, row_number):
                continue

            group_value = _optional_text(sheet.cell(row_number, 1).value)
            if group_value is not None:
                vehicle_driver_group = group_value

            destination_alias = _destination_alias(
                sheet.cell(row_number, 2).value,
                f"{sheet.title} row {row_number} destination",
            )
            day_values = [
                sheet.cell(row_number, column).value for column in range(3, 34)
            ]
            day_formula_values = [
                formula_sheet.cell(row_number, column).value
                for column in range(3, 34)
            ]
            has_missing_day_formula_cache = any(
                not _has_value(cached_value) and _is_formula(formula_value)
                for cached_value, formula_value in zip(
                    day_values, day_formula_values, strict=True
                )
            )
            cached_count_value = sheet.cell(row_number, 34).value
            unit_value = sheet.cell(row_number, 36).value
            cached_subtotal_value = sheet.cell(row_number, 37).value
            footer_marker_value = sheet.cell(row_number, 38).value
            if _is_regular_footer_or_note(
                destination_alias,
                day_values,
                cached_count_value,
                unit_value,
                cached_subtotal_value,
                footer_marker_value,
                has_missing_day_formula_cache,
            ):
                continue
            detail_values = (
                destination_alias,
                *day_values,
                cached_count_value,
                unit_value,
                cached_subtotal_value,
                footer_marker_value,
            )
            if (
                not any(_has_value(value) for value in detail_values)
                and not has_missing_day_formula_cache
            ):
                continue
            if destination_alias is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} destination is required"
                )
            for column, (cached_value, formula_value) in enumerate(
                zip(day_values, day_formula_values, strict=True), start=3
            ):
                if not _has_value(cached_value) and _is_formula(formula_value):
                    _required_cached_value(
                        cached_value,
                        formula_value,
                        sheet.title,
                        row_number,
                        formula_sheet.cell(row_number, column).coordinate,
                        "daily trip count",
                    )
            unit_formula_value = formula_sheet.cell(row_number, 36).value
            if not _has_value(unit_value) and _is_formula(unit_formula_value):
                _required_cached_value(
                    unit_value,
                    unit_formula_value,
                    sheet.title,
                    row_number,
                    f"AJ{row_number}",
                    "unit cost",
                )
            count_formula_value = formula_sheet.cell(row_number, 34).value
            if not _has_value(cached_count_value) and _is_formula(
                count_formula_value
            ):
                _required_cached_value(
                    cached_count_value,
                    count_formula_value,
                    sheet.title,
                    row_number,
                    f"AH{row_number}",
                    "cached total count",
                )
            subtotal_formula_value = formula_sheet.cell(row_number, 37).value
            if not _has_value(cached_subtotal_value) and _is_formula(
                subtotal_formula_value
            ):
                _required_cached_value(
                    cached_subtotal_value,
                    subtotal_formula_value,
                    sheet.title,
                    row_number,
                    f"AK{row_number}",
                    "cached subtotal",
                )
            if (
                not _has_value(unit_value)
                and all(_is_blank_or_numeric_zero(value) for value in day_values)
                and _is_blank_or_numeric_zero(cached_count_value)
                and _is_blank_or_numeric_zero(cached_subtotal_value)
            ):
                continue
            unit_value = _required_cached_value(
                unit_value,
                unit_formula_value,
                sheet.title,
                row_number,
                f"AJ{row_number}",
                "unit cost",
            )
            if isinstance(unit_value, bool) or not isinstance(
                unit_value, (int, float, Decimal)
            ):
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} unit cost must be numeric"
                )

            unit_rate = _integer_won(
                unit_value,
                f"{sheet.title} row {row_number} unit cost",
            )
            row_entries: list[ParsedTransportEntry] = []
            calculated_count = Decimal(0)
            for day, raw_count in enumerate(day_values, start=1):
                if raw_count in (None, ""):
                    continue
                trip_count = _trip_count_decimal(
                    raw_count,
                    f"{sheet.title} row {row_number} day {day} trip count",
                )
                if trip_count == 0:
                    continue
                source_date = _validate_calendar_day(
                    report_month, day, sheet.title, row_number
                )
                calculated_count += trip_count
                cost_won = _integer_won(
                    trip_count * unit_rate,
                    f"{sheet.title} row {row_number} day {day} calculated cost",
                )
                row_entries.append(
                    ParsedTransportEntry(
                        report_month=report_month,
                        destination_alias=destination_alias,
                        source_sheet=sheet.title,
                        source_row=row_number,
                        source_date=source_date,
                        day=day,
                        transport_type="regular",
                        trip_count=trip_count,
                        unit_rate_won=unit_rate,
                        cost_won=cost_won,
                        vehicle_driver_group=vehicle_driver_group,
                    )
                )

            cached_count = _nonnegative_decimal(
                _required_cached_value(
                    cached_count_value,
                    count_formula_value,
                    sheet.title,
                    row_number,
                    f"AH{row_number}",
                    "cached total count",
                ),
                f"{sheet.title} row {row_number} cached total count",
            )
            if calculated_count != cached_count:
                raise TotalCountMismatchError(
                    f"Total count mismatch in {sheet.title} row {row_number}: "
                    f"calculated {calculated_count}, cached {cached_count}"
                )
            cached_subtotal = _integer_won(
                _required_cached_value(
                    cached_subtotal_value,
                    subtotal_formula_value,
                    sheet.title,
                    row_number,
                    f"AK{row_number}",
                    "cached subtotal",
                ),
                f"{sheet.title} row {row_number} cached subtotal",
            )
            calculated_subtotal = sum(entry.cost_won for entry in row_entries)
            if calculated_subtotal != cached_subtotal:
                raise SubtotalMismatchError(
                    f"Subtotal mismatch in {sheet.title} row {row_number}: "
                    f"calculated {calculated_subtotal}, cached {cached_subtotal}"
                )
            rows.extend(row_entries)
        return rows

    def _parse_subcontract(
        self, sheet, formula_sheet, report_month: str
    ) -> list[ParsedTransportEntry]:
        header = _find_subcontract_header(sheet)
        if header is None:
            raise WorkbookStructureError(
                f"Required header was not found in sheet {sheet.title}"
            )
        (
            header_row,
            date_column,
            destination_column,
            tonnage_column,
            number_column,
            last_footprint_column,
        ) = header

        rows: list[ParsedTransportEntry] = []
        calculated_total = 0
        displayed_total_found = False
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if _is_displayed_total_row(sheet, row_number):
                cached_total = _integer_won(
                    _required_cached_value(
                        sheet.cell(row_number, 9).value,
                        formula_sheet.cell(row_number, 9).value,
                        sheet.title,
                        row_number,
                        f"I{row_number}",
                        "displayed cached total",
                    ),
                    f"{sheet.title} row {row_number} cached total",
                )
                if cached_total != calculated_total:
                    raise SubcontractTotalMismatchError(
                        f"Total mismatch in {sheet.title} row {row_number}: "
                        f"calculated {calculated_total}, cached {cached_total}"
                    )
                displayed_total_found = True
                break
            footprint_values = [
                sheet.cell(row_number, column).value
                for column in range(1, last_footprint_column + 1)
            ]
            if not any(_has_value(value) for value in footprint_values):
                continue
            if _is_subcontract_template_row(
                sheet, row_number, number_column, last_footprint_column
            ):
                continue
            raw_date = sheet.cell(row_number, date_column).value
            transport_date = _transport_date(
                raw_date, f"{sheet.title} row {row_number} date"
            )
            destination_alias = _destination_alias(
                sheet.cell(row_number, destination_column).value,
                f"{sheet.title} row {row_number} destination",
            )
            raw_tonnage = sheet.cell(row_number, tonnage_column).value
            raw_amount = sheet.cell(row_number, 9).value
            if destination_alias is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} destination is required"
                )
            if not _is_in_subcontract_billing_window(transport_date, report_month):
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} date is outside the "
                    f"{report_month} billing window"
                )
            vehicle_type = _optional_text(raw_tonnage)
            if vehicle_type is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} vehicle is required"
                )
            amount = _integer_won(
                _required_cached_value(
                    raw_amount,
                    formula_sheet.cell(row_number, 9).value,
                    sheet.title,
                    row_number,
                    f"I{row_number}",
                    "amount",
                ),
                f"{sheet.title} row {row_number} amount",
            )
            calculated_total += amount
            rows.append(
                ParsedTransportEntry(
                    report_month=report_month,
                    destination_alias=destination_alias,
                    source_sheet=sheet.title,
                    source_row=row_number,
                    source_date=transport_date,
                    day=transport_date.day,
                    transport_type="nonregular",
                    trip_count=Decimal(1),
                    unit_rate_won=None,
                    cost_won=amount,
                    vehicle_type=vehicle_type,
                )
            )
        if not displayed_total_found:
            raise WorkbookStructureError(
                f"{sheet.title} displayed cached total is required"
            )
        return rows


def _find_regular_header(sheet) -> int | None:
    for row_number in range(1, min(sheet.max_row, 50) + 1):
        if _is_regular_header(sheet, row_number):
            return row_number
    return None


def _is_regular_header(sheet, row_number: int) -> bool:
    destination = _header_text(sheet.cell(row_number, 2).value)
    total_count = _header_text(sheet.cell(row_number, 34).value)
    unit = _header_text(sheet.cell(row_number, 36).value)
    subtotal = _header_text(sheet.cell(row_number, 37).value)
    day_values = [sheet.cell(row_number, column).value for column in range(3, 34)]
    return (
        destination in _REGULAR_COLUMN_B_HEADERS
        and total_count in _REGULAR_TOTAL_COUNT_HEADERS
        and unit in _REGULAR_UNIT_HEADERS
        and subtotal in _REGULAR_SUBTOTAL_HEADERS
        and day_values == list(range(1, 32))
    )


def _find_subcontract_header(
    sheet,
) -> tuple[int, int, int, int, int | None, int] | None:
    for row_number in range(1, min(sheet.max_row, 50) + 1):
        columns: dict[str, int] = {}
        for column in range(1, sheet.max_column + 1):
            value = _header_text(sheet.cell(row_number, column).value)
            if value in _DATE_HEADERS:
                columns.setdefault("date", column)
            if value in _DESTINATION_HEADERS:
                columns.setdefault("destination", column)
            if value in _TONNAGE_HEADERS:
                columns.setdefault("tonnage", column)
            if value.casefold() in _NUMBER_HEADERS:
                columns.setdefault("number", column)
        amount_header = _header_text(sheet.cell(row_number, 9).value)
        if (
            {"date", "destination", "tonnage"} <= columns.keys()
            and amount_header in _AMOUNT_HEADERS
        ):
            return (
                row_number,
                columns["date"],
                columns["destination"],
                columns["tonnage"],
                columns.get("number"),
                _SUBCONTRACT_LAST_FOOTPRINT_COLUMN,
            )
    return None


def _header_text(value: object) -> str:
    return "" if value is None else str(value).strip().replace(" ", "")


def _has_value(value: object) -> bool:
    return value is not None and value != ""


def _is_blank_or_numeric_zero(value: object) -> bool:
    if not _has_value(value):
        return True
    if isinstance(value, bool):
        return False
    try:
        numeric_value = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return False
    return numeric_value.is_finite() and numeric_value == 0


def _is_displayed_total_row(sheet, row_number: int) -> bool:
    label_columns = [
        column
        for column in _SUBCONTRACT_TOTAL_LABEL_COLUMNS
        if _header_text(sheet.cell(row_number, column).value) in _TOTAL_LABELS
    ]
    if len(label_columns) != 1:
        return False
    allowed_columns = {label_columns[0], 9}
    return not any(
        _has_value(sheet.cell(row_number, column).value)
        for column in range(1, _SUBCONTRACT_LAST_FOOTPRINT_COLUMN + 1)
        if column not in allowed_columns
    )


def _is_regular_footer_or_note(
    destination_alias: str | None,
    day_values: list[object],
    cached_count_value: object,
    unit_value: object,
    cached_subtotal_value: object,
    footer_marker_value: object,
    has_missing_day_formula_cache: bool,
) -> bool:
    return (
        destination_alias is None
        and not any(_has_value(value) for value in day_values)
        and not _has_value(cached_count_value)
        and not _has_value(unit_value)
        and not has_missing_day_formula_cache
        and isinstance(cached_subtotal_value, str)
        and bool(cached_subtotal_value.strip())
        and not isinstance(footer_marker_value, bool)
        and isinstance(footer_marker_value, (int, float, Decimal))
    )


def _is_subcontract_template_row(
    sheet,
    row_number: int,
    number_column: int | None,
    last_footprint_column: int,
) -> bool:
    if number_column is None:
        return False
    return _has_value(sheet.cell(row_number, number_column).value) and not any(
        _has_value(sheet.cell(row_number, column).value)
        for column in range(1, last_footprint_column + 1)
        if column != number_column
    )


def _required_cached_value(
    cached_value: object,
    formula_value: object,
    sheet_name: str,
    row_number: int,
    coordinate: str,
    label: str,
) -> object:
    if _has_value(cached_value):
        return cached_value
    if isinstance(formula_value, str) and formula_value.startswith("="):
        raise WorkbookStructureError(
            f"{sheet_name} row {row_number} cell {coordinate} cached value is "
            "missing; recalculate and save in Excel"
        )
    raise WorkbookStructureError(
        f"{sheet_name} row {row_number} {label} is required"
    )


def _is_formula(value: object) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _optional_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _destination_alias(value: object, label: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise WorkbookStructureError(f"{label} must be text")
    text = value.strip()
    return text or None


def _validate_regular_sheet_month(sheet_name: str, report_month: str) -> None:
    match = _REGULAR_SHEET.fullmatch(sheet_name)
    if match is None:
        raise WorkbookStructureError(f"Invalid regular sheet name: {sheet_name}")
    if int(match.group("month")) != int(report_month[-2:]):
        raise WorkbookStructureError(
            f"Regular sheet month does not match report_month {report_month}"
        )


def _validate_calendar_day(
    report_month: str, day: int, sheet_name: str, row_number: int
) -> date:
    year, month = (int(part) for part in report_month.split("-"))
    try:
        return date(year, month, day)
    except ValueError as error:
        raise WorkbookStructureError(
            f"{sheet_name} row {row_number} has invalid calendar date "
            f"{report_month}-{day:02d}"
        ) from error


def _nonnegative_decimal(value: object, label: str) -> Decimal:
    if isinstance(value, bool):
        raise WorkbookStructureError(f"{label} must be a nonnegative number")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise WorkbookStructureError(f"{label} must be a nonnegative number") from error
    if not result.is_finite() or result < 0:
        raise WorkbookStructureError(f"{label} must be a nonnegative number")
    return result


def _trip_count_decimal(value: object, label: str) -> Decimal:
    if isinstance(value, bool):
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    raw_text = str(value).strip()
    if "e" in raw_text.casefold():
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    try:
        result = Decimal(raw_text)
    except (InvalidOperation, ValueError) as error:
        raise WorkbookStructureError(
            f"{label} must be a bounded nonnegative number"
        ) from error
    if (
        not result.is_finite()
        or result < 0
        or result > MAX_TRIP_COUNT
        or result.as_tuple().exponent < -MAX_TRIP_COUNT_SCALE
    ):
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    canonical_trip_count(result, label)
    return result


def canonical_trip_count(value: Decimal, label: str = "trip count") -> str:
    result = _trip_count_decimal_without_canonical_check(value, label)
    if result == 0:
        return "0"
    text = format(result, "f")
    canonical = text.rstrip("0").rstrip(".") if "." in text else text
    if len(canonical) > MAX_TRIP_COUNT_CANONICAL_LENGTH:
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    return canonical


def _trip_count_decimal_without_canonical_check(
    value: Decimal, label: str
) -> Decimal:
    if not isinstance(value, Decimal):
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    if (
        not value.is_finite()
        or value < 0
        or value > MAX_TRIP_COUNT
        or value.as_tuple().exponent < -MAX_TRIP_COUNT_SCALE
    ):
        raise WorkbookStructureError(f"{label} must be a bounded nonnegative number")
    return value


def _integer_won(value: object, label: str) -> int:
    if value is None or isinstance(value, bool):
        raise WorkbookStructureError(f"{label} must be an integer won amount")
    try:
        decimal_value = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise WorkbookStructureError(f"{label} must be an integer won amount") from error
    if (
        not decimal_value.is_finite()
        or decimal_value < 0
        or decimal_value != decimal_value.to_integral_value()
        or decimal_value > _SQLITE_MAX
    ):
        raise WorkbookStructureError(f"{label} must be an integer won amount")
    return int(decimal_value)


def _transport_date(value: object, label: str) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        normalized = value.strip().replace(".", "-").replace("/", "-")
        try:
            return date.fromisoformat(normalized.rstrip("-"))
        except ValueError as error:
            raise WorkbookStructureError(f"{label} must be a valid date") from error
    raise WorkbookStructureError(f"{label} must be a valid date")


def _is_in_subcontract_billing_window(
    transport_date: date, report_month: str
) -> bool:
    year, month = (int(part) for part in report_month.split("-"))
    report_month_start = date(year, month, 1)
    previous_month_day = report_month_start - timedelta(days=1)
    return (transport_date.year, transport_date.month) in {
        (year, month),
        (previous_month_day.year, previous_month_day.month),
    }
