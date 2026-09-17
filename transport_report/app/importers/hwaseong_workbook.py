from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
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
_TOTAL_LABELS = {"합계", "총계", "계"}
_SQLITE_MAX = 2**63 - 1


class TransportWorkbookError(Exception):
    """Base error for workbooks that cannot be imported safely."""


class WorkbookStructureError(TransportWorkbookError):
    """Raised when a required workbook structure or value is missing."""


class SubtotalMismatchError(TransportWorkbookError):
    """Raised when a regular row does not reconcile to its cached subtotal."""


class TotalCountMismatchError(TransportWorkbookError):
    """Raised when a regular row does not reconcile to its cached trip count."""


@dataclass(frozen=True, slots=True)
class ParsedTransportEntry:
    report_month: str
    destination_alias: str
    source_sheet: str
    source_row: int
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
        try:
            workbook = load_workbook(
                path,
                data_only=True,
                read_only=False,
                keep_vba=False,
                keep_links=False,
            )
        except Exception as error:
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
            rows = self._parse_regular(regular_sheet, report_month)
            for sheet_name in SUBCONTRACT_SHEETS:
                rows.extend(
                    self._parse_subcontract(workbook[sheet_name], report_month)
                )
            return rows
        finally:
            workbook.close()

    def _parse_regular(self, sheet, report_month: str) -> list[ParsedTransportEntry]:
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
            cached_count_value = sheet.cell(row_number, 34).value
            unit_value = sheet.cell(row_number, 36).value
            cached_subtotal_value = sheet.cell(row_number, 37).value
            detail_values = (
                destination_alias,
                *day_values,
                cached_count_value,
                unit_value,
                cached_subtotal_value,
            )
            if not any(_has_value(value) for value in detail_values):
                continue
            if destination_alias is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} destination is required"
                )
            if (
                not _has_value(unit_value)
                and all(_is_blank_or_numeric_zero(value) for value in day_values)
                and _is_blank_or_numeric_zero(cached_count_value)
                and _is_blank_or_numeric_zero(cached_subtotal_value)
            ):
                continue
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
                trip_count = _nonnegative_decimal(
                    raw_count,
                    f"{sheet.title} row {row_number} day {day} trip count",
                )
                if trip_count == 0:
                    continue
                _validate_calendar_day(report_month, day, sheet.title, row_number)
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
                        day=day,
                        transport_type="regular",
                        trip_count=trip_count,
                        unit_rate_won=unit_rate,
                        cost_won=cost_won,
                        vehicle_driver_group=vehicle_driver_group,
                    )
                )

            cached_count = _nonnegative_decimal(
                cached_count_value,
                f"{sheet.title} row {row_number} cached total count",
            )
            if calculated_count != cached_count:
                raise TotalCountMismatchError(
                    f"Total count mismatch in {sheet.title} row {row_number}: "
                    f"calculated {calculated_count}, cached {cached_count}"
                )
            cached_subtotal = _integer_won(
                cached_subtotal_value,
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
        self, sheet, report_month: str
    ) -> list[ParsedTransportEntry]:
        header = _find_subcontract_header(sheet)
        if header is None:
            raise WorkbookStructureError(
                f"Required header was not found in sheet {sheet.title}"
            )
        header_row, date_column, destination_column, tonnage_column = header

        rows: list[ParsedTransportEntry] = []
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if _is_displayed_total_row(sheet, row_number):
                break
            raw_date = sheet.cell(row_number, date_column).value
            destination_alias = _destination_alias(
                sheet.cell(row_number, destination_column).value,
                f"{sheet.title} row {row_number} destination",
            )
            raw_tonnage = sheet.cell(row_number, tonnage_column).value
            raw_amount = sheet.cell(row_number, 9).value
            if not any(
                _has_value(value)
                for value in (raw_date, destination_alias, raw_tonnage, raw_amount)
            ):
                continue
            if destination_alias is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} destination is required"
                )
            transport_date = _transport_date(
                raw_date, f"{sheet.title} row {row_number} date"
            )
            if transport_date.strftime("%Y-%m") != report_month:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} date is outside {report_month}"
                )
            amount = _integer_won(
                raw_amount, f"{sheet.title} row {row_number} amount"
            )
            rows.append(
                ParsedTransportEntry(
                    report_month=report_month,
                    destination_alias=destination_alias,
                    source_sheet=sheet.title,
                    source_row=row_number,
                    day=transport_date.day,
                    transport_type="nonregular",
                    trip_count=Decimal(1),
                    unit_rate_won=None,
                    cost_won=amount,
                    vehicle_type=_optional_text(raw_tonnage),
                )
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


def _find_subcontract_header(sheet) -> tuple[int, int, int, int] | None:
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
    return any(
        _header_text(sheet.cell(row_number, column).value) in _TOTAL_LABELS
        for column in range(1, 10)
    )


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
) -> None:
    year, month = (int(part) for part in report_month.split("-"))
    try:
        date(year, month, day)
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
