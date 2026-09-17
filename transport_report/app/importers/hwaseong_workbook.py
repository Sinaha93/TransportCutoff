from __future__ import annotations

import re
from collections import deque
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from posixpath import join as posix_join
from posixpath import normpath as posix_normpath
from xml.etree.ElementTree import ParseError, iterparse
from zipfile import BadZipFile, LargeZipFile, ZipFile

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
# Layout bounds are intentionally well above the real workbook's 116 rows/54 columns.
MAX_WORKSHEET_ROWS = 2_000
MAX_WORKSHEET_COLUMNS = 128
MAX_PARSED_ENTRIES = 20_000
# Pre-load limits cover every worksheet, not only the four imported sheets. The
# real source has 6 sheets, at most 3,970 cell records/sheet, 7,535 total cells,
# 116 rows, 54 columns, and a largest worksheet XML of 101,774 bytes.
# Expanded merged-range areas are charged against the same cell envelopes because
# OpenPyXL creates a MergedCell object for each covered coordinate.
MAX_PREFLIGHT_WORKSHEETS = 256
MAX_PREFLIGHT_WORKSHEET_XML_BYTES = 8 * 1024 * 1024
MAX_PREFLIGHT_RELATED_XML_BYTES = 8 * 1024 * 1024
MAX_PREFLIGHT_PACKAGE_XML_BYTES = 8 * 1024 * 1024
MAX_PREFLIGHT_RELATED_PART_BYTES = 64 * 1024 * 1024
# OpenPyXL 3.1.5 load audit: manifest, shared strings, workbook, core/custom
# properties, theme, styles, then worksheets and their comments/tables/drawings/
# charts/images/pivots. keep_links=False skips external-link bodies. app.xml is
# not currently materialized but is checked as package metadata. Generic package
# trees are stream-validated under 8 MiB, which bounds their one-shot parse cost;
# shared strings/styles need the deeper object/text counts below because OpenPyXL
# eagerly builds a Python object for every record. Excel's documented unique-cell-
# format ceiling is 65,490, so 65,536 leaves ordinary workbooks headroom.
MAX_PREFLIGHT_CONTENT_TYPE_RECORDS = 4_096
MAX_PREFLIGHT_SHARED_STRING_RECORDS = 100_000
# OpenPyXL creates a Text object for each <si> and a RichText/PhoneticText
# object for every direct <r>/<rPh> child. Their non-repeating descendants
# (rPr/phoneticPr and their properties) are bounded one-for-one by these parent
# records, so this keeps the eager shared-string object graph bounded as well.
MAX_PREFLIGHT_SHARED_STRING_OBJECT_RECORDS = 100_000
MAX_PREFLIGHT_SHARED_STRING_CHARACTERS = 8_000_000
MAX_PREFLIGHT_STYLE_COLLECTION_RECORDS = 65_536
MAX_PREFLIGHT_STYLE_RECORDS = 100_000
MAX_PREFLIGHT_RELATIONSHIP_PARTS = 1_024
# A relationship part is parsed into OpenPyXL relationship objects even when its
# records are unused. Bound records before materializing them; the four-part
# global allowance also prevents many smaller sidecars from bypassing the cap.
MAX_PREFLIGHT_RELATIONSHIP_RECORDS = 4_096
MAX_PREFLIGHT_TOTAL_RELATIONSHIP_RECORDS = 16_384
MAX_PREFLIGHT_RELATIONSHIP_EDGES = 4_096
MAX_PREFLIGHT_RELATIONSHIP_DEPTH = 64
MAX_PREFLIGHT_RELATIONSHIP_QUEUE = 1_024
MAX_PREFLIGHT_HYPERLINK_TARGET_LENGTH = 8_192
MAX_PREFLIGHT_WORKSHEET_CELLS = 50_000
MAX_PREFLIGHT_TOTAL_CELLS = 100_000
MAX_PREFLIGHT_ROWS = 20_000
MAX_PREFLIGHT_COLUMNS = 512
# OpenPyXL creates one ColumnDimension per raw <col> record; it does not expand
# min..max spans. The real source has up to 704 records and valid style spans
# through Excel's final column, so bound records separately from data columns.
MAX_PREFLIGHT_COLUMN_DIMENSION_RECORDS = 2_048
MAX_EXCEL_COLUMNS = 16_384
# Conservative source-cell limits: 10,000 trips/day, 4 decimals, 16 canonical chars.
MAX_TRIP_COUNT = Decimal("10000")
MAX_TRIP_COUNT_SCALE = 4
MAX_TRIP_COUNT_CANONICAL_LENGTH = 16
_CELL_REFERENCE = re.compile(r"^(?P<column>[A-Za-z]+)(?P<row>[1-9][0-9]*)$")
_DOCUMENT_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_STRICT_DOCUMENT_RELATIONSHIP_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships"
)
_RELATIONSHIP_SEMANTIC_ALIASES = {
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle": (
        "chartStyle"
    ),
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle": (
        "chartColorStyle"
    ),
}
_DOCUMENT_RELATIONSHIP_ID = f"{{{_DOCUMENT_RELATIONSHIP_NAMESPACE}}}id"
_DOCUMENT_RELATIONSHIP_REFERENCES = {
    f"{{{_DOCUMENT_RELATIONSHIP_NAMESPACE}}}{name}"
    for name in ("id", "embed", "link")
}
_HYPERLINK_RELATIONSHIPS = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink",
}
_HYPERLINK_CONTROL_CHARACTERS = re.compile(r"[\x00-\x1f\x7f-\x9f]")
_CONTENT_TYPE = re.compile(r"^[!#$&^_.+A-Za-z0-9-]+/[!#$&^_.+A-Za-z0-9-]+$")
_CONTENT_TYPES_ROOT = (
    "{http://schemas.openxmlformats.org/package/2006/content-types}Types"
)
_SPREADSHEETML_MAIN_NAMESPACE = (
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)
_SHARED_STRING_ITEM_TAG = f"{{{_SPREADSHEETML_MAIN_NAMESPACE}}}si"
_SHARED_STRING_RICH_TEXT_TAGS = {
    f"{{{_SPREADSHEETML_MAIN_NAMESPACE}}}r",
    f"{{{_SPREADSHEETML_MAIN_NAMESPACE}}}rPh",
}

_CONTENT_TYPE_WORKBOOKS = (
    "application/vnd.ms-excel.template.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
)
_CONTENT_TYPE_SHARED_STRINGS = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)
_CONTENT_TYPE_STYLES = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
)
_CONTENT_TYPE_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
_CONTENT_TYPE_CORE_PROPERTIES = (
    "application/vnd.openxmlformats-package.core-properties+xml"
)
_CONTENT_TYPE_EXTENDED_PROPERTIES = (
    "application/vnd.openxmlformats-officedocument.extended-properties+xml"
)
_CONTENT_TYPE_CUSTOM_PROPERTIES = (
    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
)

_XML_RELATIONSHIP_CONTENT_TYPES = {
    "worksheet": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    },
    "dialogsheet": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    },
    "macrosheet": {
        "application/vnd.ms-excel.macrosheet+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    },
    "intlMacrosheet": {
        "application/vnd.ms-excel.intlmacrosheet+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    },
    "chartsheet": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
    },
    "drawing": {
        "application/vnd.openxmlformats-officedocument.drawing+xml"
    },
    "chart": {
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    },
    "table": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    },
    "queryTable": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"
    },
    "pivotTable": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    },
    "pivotCacheDefinition": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
    },
    "pivotCacheRecords": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
    },
    "comments": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    },
    "vmlDrawing": {
        "application/vnd.openxmlformats-officedocument.vmlDrawing"
    },
    "chartStyle": {"application/vnd.ms-office.chartstyle+xml"},
    "chartColorStyle": {"application/vnd.ms-office.chartcolorstyle+xml"},
    "styles": {_CONTENT_TYPE_STYLES},
    "theme": {_CONTENT_TYPE_THEME},
    "sharedStrings": {_CONTENT_TYPE_SHARED_STRINGS},
}
_BINARY_RELATIONSHIP_CONTENT_TYPE_PREFIXES = {
    "image": ("image/",),
    "audio": ("audio/",),
    "video": ("video/",),
}
_STYLE_COLLECTION_ELEMENTS = {
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "indexedColors",
}


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


@dataclass(frozen=True, slots=True)
class _Relationship:
    id: str
    type: str
    target: str
    target_mode: str | None


@dataclass(frozen=True, slots=True)
class _RelationshipTarget:
    member_name: str
    relationship_id: str
    relationship_type: str
    source_member: str
    expected_semantic: str | None = None


class _RelationshipRecordBudget:
    def __init__(self) -> None:
        self._counted_members: set[str] = set()
        self._total_records = 0

    def begins_part(self, member_name: str) -> bool:
        if member_name in self._counted_members:
            return False
        self._counted_members.add(member_name)
        return True

    def record(
        self,
        description: str,
        part_records: int,
        count_toward_total: bool,
    ) -> None:
        if part_records > MAX_PREFLIGHT_RELATIONSHIP_RECORDS:
            raise WorkbookStructureError(
                f"{description} exceeds relationship record limit of "
                f"{MAX_PREFLIGHT_RELATIONSHIP_RECORDS}"
            )
        if not count_toward_total:
            return
        self._total_records += 1
        if self._total_records > MAX_PREFLIGHT_TOTAL_RELATIONSHIP_RECORDS:
            raise WorkbookStructureError(
                "Workbook exceeds total relationship record limit of "
                f"{MAX_PREFLIGHT_TOTAL_RELATIONSHIP_RECORDS}"
            )


@dataclass(frozen=True, slots=True)
class _ContentTypes:
    defaults: dict[str, str]
    overrides: dict[str, str]
    overrides_in_order: tuple[tuple[str, str], ...]

    def for_member(self, member_name: str) -> str | None:
        override = self.overrides.get(member_name)
        if override is not None:
            return override
        filename = PurePosixPath(member_name).name
        if "." not in filename:
            return None
        return self.defaults.get(filename.rsplit(".", 1)[1].lower())

    def members_with_type(self, content_type: str) -> list[str]:
        return [
            member_name
            for member_name, mapped_type in self.overrides_in_order
            if mapped_type == content_type
        ]


class HwaseongWorkbookParser:
    def parse(
        self, workbook_path: str | Path, report_month: str
    ) -> list[ParsedTransportEntry]:
        if not isinstance(report_month, str) or _MONTH.fullmatch(report_month) is None:
            raise WorkbookStructureError("report_month must use YYYY-MM")

        path = Path(workbook_path)
        _preflight_workbook_resources(path)
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
            for sheet_name in (regular_sheets[0], *SUBCONTRACT_SHEETS):
                _validate_worksheet_bounds(workbook[sheet_name])
            rows = self._parse_regular(
                regular_sheet,
                formula_workbook[regular_sheets[0]],
                report_month,
                MAX_PARSED_ENTRIES,
            )
            for sheet_name in SUBCONTRACT_SHEETS:
                rows.extend(
                    self._parse_subcontract(
                        workbook[sheet_name],
                        formula_workbook[sheet_name],
                        report_month,
                        MAX_PARSED_ENTRIES - len(rows),
                    )
                )
            return rows
        finally:
            workbook.close()
            formula_workbook.close()

    def _parse_regular(
        self, sheet, formula_sheet, report_month: str, entry_limit: int
    ) -> list[ParsedTransportEntry]:
        header_row = _find_regular_header(sheet, formula_sheet)
        if header_row is None:
            raise WorkbookStructureError(
                f"Required header was not found in sheet {sheet.title}"
            )

        rows: list[ParsedTransportEntry] = []
        vehicle_driver_group: str | None = None
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if _is_regular_header(sheet, formula_sheet, row_number):
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
            cached_count_value = sheet.cell(row_number, 34).value
            unit_value = sheet.cell(row_number, 36).value
            cached_subtotal_value = sheet.cell(row_number, 37).value
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
            if _has_regular_footer_label(sheet, formula_sheet, row_number):
                cached_footer_marker = sheet.cell(row_number, 38).value
                formula_footer_marker = formula_sheet.cell(row_number, 38).value
                if not _has_value(cached_footer_marker) and _is_formula(
                    formula_footer_marker
                ):
                    _required_cached_value(
                        cached_footer_marker,
                        formula_footer_marker,
                        sheet.title,
                        row_number,
                        f"AL{row_number}",
                        "footer marker",
                    )
            if _is_regular_footer_or_note(sheet, formula_sheet, row_number):
                continue
            if _is_regular_blank_or_group_row(sheet, formula_sheet, row_number):
                continue
            if destination_alias is None:
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} destination is required"
                )
            if _is_regular_inactive_row(sheet, formula_sheet, row_number):
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
                _append_parsed_entry(
                    row_entries,
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
                    ),
                    entry_limit - len(rows),
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
        self, sheet, formula_sheet, report_month: str, entry_limit: int
    ) -> list[ParsedTransportEntry]:
        header = _find_subcontract_header(sheet, formula_sheet)
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
            if _is_displayed_total_row(sheet, formula_sheet, row_number):
                if displayed_total_found:
                    raise WorkbookStructureError(
                        f"{sheet.title} row {row_number} contains a second displayed total"
                    )
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
                continue
            footprint_values = [
                sheet.cell(row_number, column).value
                for column in range(1, last_footprint_column + 1)
            ]
            formula_footprint_values = [
                formula_sheet.cell(row_number, column).value
                for column in range(1, last_footprint_column + 1)
            ]
            if not any(
                _has_value(value)
                for value in (*footprint_values, *formula_footprint_values)
            ):
                continue
            if _is_subcontract_template_row(
                sheet,
                formula_sheet,
                row_number,
                number_column,
                last_footprint_column,
            ):
                continue
            if displayed_total_found:
                if _is_subcontract_formula_footer(
                    sheet, formula_sheet, row_number, last_footprint_column
                ):
                    continue
                raise WorkbookStructureError(
                    f"{sheet.title} row {row_number} contains detail after "
                    "the displayed total"
                )
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
            _append_parsed_entry(
                rows,
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
                ),
                entry_limit,
            )
        if not displayed_total_found:
            raise WorkbookStructureError(
                f"{sheet.title} displayed cached total is required"
            )
        return rows


def _validate_worksheet_bounds(sheet) -> None:
    if sheet.max_row > MAX_WORKSHEET_ROWS:
        raise WorkbookStructureError(
            f"{sheet.title} exceeds worksheet row limit of {MAX_WORKSHEET_ROWS} "
            f"(found {sheet.max_row})"
        )
    if sheet.max_column > MAX_WORKSHEET_COLUMNS:
        raise WorkbookStructureError(
            f"{sheet.title} exceeds worksheet column limit of "
            f"{MAX_WORKSHEET_COLUMNS} (found {sheet.max_column})"
        )


def _read_content_types(archive: ZipFile) -> _ContentTypes:
    try:
        member = archive.getinfo("[Content_Types].xml")
    except KeyError as error:
        raise WorkbookStructureError("Content types XML is missing") from error
    if member.is_dir():
        raise WorkbookStructureError("Content types XML is invalid")
    if member.file_size > MAX_PREFLIGHT_PACKAGE_XML_BYTES:
        raise WorkbookStructureError(
            "Content types XML exceeds XML size limit of "
            f"{MAX_PREFLIGHT_PACKAGE_XML_BYTES}"
        )

    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    overrides_in_order: list[tuple[str, str]] = []
    record_count = 0
    root_seen = False
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_PACKAGE_XML_BYTES,
                "Content types XML",
            )
            for event, element in iterparse(
                limited_source, events=("start", "end")
            ):
                if not root_seen:
                    root_seen = True
                    if event != "start" or element.tag != _CONTENT_TYPES_ROOT:
                        raise WorkbookStructureError(
                            "Content types XML is invalid"
                        )
                if event == "start":
                    continue
                local_name = _xml_local_name(element.tag)
                if local_name == "Default":
                    record_count += 1
                    extension = element.attrib.get("Extension", "")
                    content_type = element.attrib.get("ContentType", "")
                    if (
                        not extension
                        or extension.startswith(".")
                        or not extension.isascii()
                        or not extension.replace("-", "").replace("_", "").isalnum()
                    ):
                        raise WorkbookStructureError(
                            "Content types XML contains an unsafe default extension"
                        )
                    _validate_content_type(content_type)
                    extension = extension.lower()
                    previous = defaults.get(extension)
                    if previous is not None and previous != content_type:
                        raise WorkbookStructureError(
                            "Content types XML contains a conflicting default"
                        )
                    defaults[extension] = content_type
                elif local_name == "Override":
                    record_count += 1
                    member_name = _safe_content_type_part_name(
                        element.attrib.get("PartName", "")
                    )
                    content_type = element.attrib.get("ContentType", "")
                    _validate_content_type(content_type)
                    previous = overrides.get(member_name)
                    if previous is not None and previous != content_type:
                        raise WorkbookStructureError(
                            "Content types XML contains a conflicting content type "
                            f"override for {member_name}"
                        )
                    if previous is None:
                        overrides[member_name] = content_type
                        overrides_in_order.append((member_name, content_type))
                if record_count > MAX_PREFLIGHT_CONTENT_TYPE_RECORDS:
                    raise WorkbookStructureError(
                        "Workbook exceeds content type record limit of "
                        f"{MAX_PREFLIGHT_CONTENT_TYPE_RECORDS}"
                    )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError("Content types XML is malformed") from error
    if not root_seen:
        raise WorkbookStructureError("Content types XML is invalid")
    return _ContentTypes(defaults, overrides, tuple(overrides_in_order))


def _safe_content_type_part_name(part_name: str) -> str:
    if (
        not part_name.startswith("/")
        or part_name.startswith("//")
        or part_name.endswith("/")
        or "\\" in part_name
        or "%" in part_name
        or "?" in part_name
        or "#" in part_name
        or ":" in part_name
        or _HYPERLINK_CONTROL_CHARACTERS.search(part_name)
    ):
        raise WorkbookStructureError(
            "Content types XML contains an unsafe part name"
        )
    parts = part_name[1:].split("/")
    if not parts or any(part in ("", ".", "..") for part in parts):
        raise WorkbookStructureError(
            "Content types XML contains an unsafe part name"
        )
    return "/".join(parts)


def _validate_content_type(content_type: str) -> None:
    if len(content_type) > 255 or _CONTENT_TYPE.fullmatch(content_type) is None:
        raise WorkbookStructureError(
            "Content types XML contains an unsafe content type"
        )


def _select_workbook_member(content_types: _ContentTypes) -> str:
    candidates: list[str] = []
    for content_type in _CONTENT_TYPE_WORKBOOKS:
        candidates.extend(content_types.members_with_type(content_type))
        if candidates:
            break
    if not candidates:
        default_types = set(content_types.defaults.values())
        if default_types.intersection(_CONTENT_TYPE_WORKBOOKS):
            candidates = ["xl/workbook.xml"]
    if len(candidates) != 1:
        raise WorkbookStructureError(
            "Content types XML must identify exactly one workbook part"
        )
    workbook_member = candidates[0]
    if workbook_member != "xl/workbook.xml":
        raise WorkbookStructureError(
            "Workbook part must be xl/workbook.xml for safe import"
        )
    return workbook_member


def _require_content_type(
    content_types: _ContentTypes,
    member_name: str,
    expected: set[str] | tuple[str, ...] | str,
    description: str,
) -> str:
    content_type = content_types.for_member(member_name)
    if content_type is None:
        raise WorkbookStructureError(
            f"{member_name} content type mapping is missing for {description}"
        )
    expected_types = {expected} if isinstance(expected, str) else set(expected)
    if content_type not in expected_types:
        raise WorkbookStructureError(
            f"{member_name} content type conflicts with {description} semantics"
        )
    return content_type


def _preflight_package_parts(
    archive: ZipFile,
    content_types: _ContentTypes,
    workbook_member: str,
    relationship_records: _RelationshipRecordBudget,
) -> None:
    _require_content_type(
        content_types, workbook_member, _CONTENT_TYPE_WORKBOOKS, "workbook"
    )
    _preflight_bounded_xml(archive, workbook_member, "Workbook")

    shared_string_members = content_types.members_with_type(
        _CONTENT_TYPE_SHARED_STRINGS
    )
    if len(shared_string_members) > 1:
        raise WorkbookStructureError(
            "Content types XML identifies multiple shared string parts"
        )
    if shared_string_members:
        _preflight_shared_strings(archive, shared_string_members[0])

    package_parts = (
        ("xl/styles.xml", _CONTENT_TYPE_STYLES, "Styles", _preflight_styles),
        ("xl/theme/theme1.xml", _CONTENT_TYPE_THEME, "Theme", None),
        (
            "docProps/core.xml",
            _CONTENT_TYPE_CORE_PROPERTIES,
            "Core properties",
            None,
        ),
        (
            "docProps/app.xml",
            _CONTENT_TYPE_EXTENDED_PROPERTIES,
            "Extended properties",
            None,
        ),
        (
            "docProps/custom.xml",
            _CONTENT_TYPE_CUSTOM_PROPERTIES,
            "Custom properties",
            None,
        ),
    )
    archive_names = set(archive.namelist())
    for member_name, content_type, description, specialized_preflight in package_parts:
        if member_name not in archive_names:
            continue
        _require_content_type(
            content_types, member_name, content_type, description.lower()
        )
        if specialized_preflight is None:
            _preflight_bounded_xml(archive, member_name, description)
        else:
            specialized_preflight(archive, member_name)
    _preflight_package_relationships(
        archive, content_types, workbook_member, relationship_records
    )


def _preflight_package_relationships(
    archive: ZipFile,
    content_types: _ContentTypes,
    workbook_member: str,
    relationship_records: _RelationshipRecordBudget,
) -> None:
    root_relationships = _read_package_relationships(
        archive, "_rels/.rels", "Package relationships", relationship_records
    )
    office_document_count = 0
    for relationship in root_relationships:
        semantic = _relationship_semantic(relationship.type)
        expected_member: str | None = None
        expected_type: str | tuple[str, ...] | None = None
        description = semantic or relationship.type
        if semantic == "officeDocument":
            office_document_count += 1
            expected_member = workbook_member
            expected_type = _CONTENT_TYPE_WORKBOOKS
        elif relationship.type == (
            "http://schemas.openxmlformats.org/package/2006/relationships/"
            "metadata/core-properties"
        ):
            expected_member = "docProps/core.xml"
            expected_type = _CONTENT_TYPE_CORE_PROPERTIES
            description = "core properties"
        elif semantic == "extended-properties":
            expected_member = "docProps/app.xml"
            expected_type = _CONTENT_TYPE_EXTENDED_PROPERTIES
        elif semantic == "custom-properties":
            expected_member = "docProps/custom.xml"
            expected_type = _CONTENT_TYPE_CUSTOM_PROPERTIES
        if expected_member is None or expected_type is None:
            continue
        if relationship.target_mode == "External":
            raise WorkbookStructureError(
                f"Package has an external {description} relationship"
            )
        target_member = _safe_package_relationship_target(
            "", relationship.target
        )
        if target_member != expected_member:
            raise WorkbookStructureError(
                f"Package {description} relationship target conflicts with "
                "the eagerly loaded part"
            )
        _required_graph_member(archive, target_member)
        _require_content_type(
            content_types, target_member, expected_type, description
        )
    if office_document_count != 1:
        raise WorkbookStructureError(
            "Package must contain exactly one office document relationship"
        )

    workbook_relationships = _read_package_relationships(
        archive,
        "xl/_rels/workbook.xml.rels",
        "Workbook relationships",
        relationship_records,
    )
    expected_members = {
        "styles": "xl/styles.xml",
        "theme": "xl/theme/theme1.xml",
    }
    shared_string_members = content_types.members_with_type(
        _CONTENT_TYPE_SHARED_STRINGS
    )
    if shared_string_members:
        expected_members["sharedStrings"] = shared_string_members[0]
    for relationship in workbook_relationships:
        semantic = _relationship_semantic(relationship.type)
        expected_member = expected_members.get(semantic or "")
        if expected_member is None:
            continue
        if relationship.target_mode == "External":
            raise WorkbookStructureError(
                f"Workbook has an external {semantic} relationship"
            )
        target_member = _safe_package_relationship_target(
            workbook_member, relationship.target
        )
        if target_member != expected_member:
            raise WorkbookStructureError(
                f"Workbook {semantic} relationship target conflicts with "
                "the eagerly loaded part"
            )
        _required_graph_member(archive, target_member)
        _validate_relationship_target_semantics(
            content_types,
            _RelationshipTarget(
                target_member,
                relationship.id,
                relationship.type,
                workbook_member,
            ),
        )


def _read_package_relationships(
    archive: ZipFile,
    member_name: str,
    description: str,
    relationship_records: _RelationshipRecordBudget,
) -> list[_Relationship]:
    member = _required_xml_member(
        archive, member_name, f"{description} XML is missing"
    )
    if member.file_size > MAX_PREFLIGHT_PACKAGE_XML_BYTES:
        raise WorkbookStructureError(
            f"{description} exceeds XML size limit of "
            f"{MAX_PREFLIGHT_PACKAGE_XML_BYTES}"
        )
    relationships: list[_Relationship] = []
    relationship_ids: set[str] = set()
    part_records = 0
    count_toward_total = relationship_records.begins_part(member_name)
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source, MAX_PREFLIGHT_PACKAGE_XML_BYTES, description
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "Relationship":
                    part_records += 1
                    relationship_records.record(
                        description, part_records, count_toward_total
                    )
                    relationship_id = element.attrib.get("Id")
                    relationship_type = element.attrib.get("Type")
                    target = element.attrib.get("Target")
                    if (
                        not relationship_id
                        or not relationship_type
                        or not target
                        or relationship_id in relationship_ids
                    ):
                        raise WorkbookStructureError(
                            f"{description} XML is invalid"
                        )
                    relationship_ids.add(relationship_id)
                    relationships.append(
                        _Relationship(
                            relationship_id,
                            relationship_type,
                            target,
                            element.attrib.get("TargetMode"),
                        )
                    )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"{description} XML is malformed"
        ) from error
    return relationships


def _safe_package_relationship_target(source_member: str, target: str) -> str:
    if (
        not target
        or target.startswith("//")
        or "\\" in target
        or "%" in target
        or "\x00" in target
        or ":" in target
        or "?" in target
        or "#" in target
    ):
        raise WorkbookStructureError(
            "Package contains an unsafe relationship target"
        )
    parts = [] if target.startswith("/") else list(
        PurePosixPath(source_member).parent.parts
    )
    if parts == ["."]:
        parts = []
    for part in PurePosixPath(target).parts:
        if part in ("/", "."):
            continue
        if part == "..":
            if not parts:
                raise WorkbookStructureError(
                    "Package contains an unsafe relationship target"
                )
            parts.pop()
            continue
        parts.append(part)
    if not parts:
        raise WorkbookStructureError(
            "Package contains an unsafe relationship target"
        )
    return "/".join(parts)


def _preflight_bounded_xml(
    archive: ZipFile, member_name: str, description: str
) -> None:
    member = _required_xml_member(
        archive, member_name, f"Package XML is missing: {member_name}"
    )
    if member.file_size > MAX_PREFLIGHT_PACKAGE_XML_BYTES:
        raise WorkbookStructureError(
            f"{description} exceeds XML size limit of "
            f"{MAX_PREFLIGHT_PACKAGE_XML_BYTES}"
        )
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source, MAX_PREFLIGHT_PACKAGE_XML_BYTES, description
            )
            for _, element in iterparse(limited_source, events=("end",)):
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Package XML is malformed: {member_name}"
        ) from error


def _preflight_shared_strings(archive: ZipFile, member_name: str) -> None:
    member = _required_xml_member(
        archive, member_name, f"Shared strings XML is missing: {member_name}"
    )
    if member.file_size > MAX_PREFLIGHT_PACKAGE_XML_BYTES:
        raise WorkbookStructureError(
            "Shared strings XML exceeds XML size limit of "
            f"{MAX_PREFLIGHT_PACKAGE_XML_BYTES}"
        )
    record_count = 0
    object_record_count = 0
    character_count = 0
    element_stack: list[str] = []
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source, MAX_PREFLIGHT_PACKAGE_XML_BYTES, "Shared strings XML"
            )
            for event, element in iterparse(limited_source, events=("start", "end")):
                if event == "start":
                    element_stack.append(element.tag)
                    continue
                local_name = _xml_local_name(element.tag)
                if local_name == "t" and element.text:
                    character_count += len(element.text)
                    if character_count > MAX_PREFLIGHT_SHARED_STRING_CHARACTERS:
                        raise WorkbookStructureError(
                            "Workbook exceeds shared string character limit of "
                            f"{MAX_PREFLIGHT_SHARED_STRING_CHARACTERS}"
                        )
                if element.tag == _SHARED_STRING_ITEM_TAG:
                    record_count += 1
                    if record_count > MAX_PREFLIGHT_SHARED_STRING_RECORDS:
                        raise WorkbookStructureError(
                            "Workbook exceeds shared string record limit of "
                            f"{MAX_PREFLIGHT_SHARED_STRING_RECORDS}"
                        )
                    object_record_count += 1
                elif (
                    len(element_stack) >= 2
                    and element_stack[-2] == _SHARED_STRING_ITEM_TAG
                    and element.tag in _SHARED_STRING_RICH_TEXT_TAGS
                ):
                    object_record_count += 1
                if object_record_count > MAX_PREFLIGHT_SHARED_STRING_OBJECT_RECORDS:
                    raise WorkbookStructureError(
                        "Workbook exceeds shared string object limit of "
                        f"{MAX_PREFLIGHT_SHARED_STRING_OBJECT_RECORDS}"
                    )
                element.clear()
                element_stack.pop()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Shared strings XML is malformed: {member_name}"
        ) from error


def _preflight_styles(archive: ZipFile, member_name: str) -> None:
    member = _required_xml_member(
        archive, member_name, f"Styles XML is missing: {member_name}"
    )
    if member.file_size > MAX_PREFLIGHT_PACKAGE_XML_BYTES:
        raise WorkbookStructureError(
            "Styles XML exceeds XML size limit of "
            f"{MAX_PREFLIGHT_PACKAGE_XML_BYTES}"
        )
    total_records = 0
    collection_counts = {name: 0 for name in _STYLE_COLLECTION_ELEMENTS}
    stack: list[str] = []
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source, MAX_PREFLIGHT_PACKAGE_XML_BYTES, "Styles XML"
            )
            for event, element in iterparse(
                limited_source, events=("start", "end")
            ):
                local_name = _xml_local_name(element.tag)
                if event == "start":
                    stack.append(local_name)
                    continue
                if local_name not in _STYLE_COLLECTION_ELEMENTS and local_name != "styleSheet":
                    total_records += 1
                    if total_records > MAX_PREFLIGHT_STYLE_RECORDS:
                        raise WorkbookStructureError(
                            "Workbook exceeds style record limit of "
                            f"{MAX_PREFLIGHT_STYLE_RECORDS}"
                        )
                if len(stack) >= 2 and stack[-2] in collection_counts:
                    collection_name = stack[-2]
                    collection_counts[collection_name] += 1
                    if (
                        collection_counts[collection_name]
                        > MAX_PREFLIGHT_STYLE_COLLECTION_RECORDS
                    ):
                        raise WorkbookStructureError(
                            f"Workbook exceeds {collection_name} style collection "
                            "record limit of "
                            f"{MAX_PREFLIGHT_STYLE_COLLECTION_RECORDS}"
                        )
                stack.pop()
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Styles XML is malformed: {member_name}"
        ) from error


def _preflight_workbook_resources(path: Path) -> None:
    try:
        with ZipFile(path, "r") as archive:
            content_types = _read_content_types(archive)
            workbook_member = _select_workbook_member(content_types)
            relationship_records = _RelationshipRecordBudget()
            _preflight_package_parts(
                archive, content_types, workbook_member, relationship_records
            )
            sheet_refs = _read_workbook_sheet_refs(archive)
            worksheet_members, chartsheet_members = _resolve_sheet_members(
                archive, sheet_refs, content_types, relationship_records
            )
            relationship_roots: dict[str, list[_RelationshipTarget]] = {}
            for _, member_name, relationship in chartsheet_members:
                _add_relationship_root(
                    relationship_roots,
                    member_name,
                    relationship,
                    workbook_member,
                    "chartsheet",
                )
            for sheet_name, member_name, _ in chartsheet_members:
                _preflight_chartsheet_xml(archive, sheet_name, member_name)
            total_cells = 0
            total_materialized_cells = 0
            for sheet_name, member_name, _ in worksheet_members:
                (
                    cell_count,
                    materialized_cell_count,
                    table_ids,
                    hyperlink_ids,
                ) = (
                    _preflight_worksheet_xml(archive, sheet_name, member_name)
                )
                comment_count = _preflight_worksheet_relationships(
                    archive,
                    sheet_name,
                    member_name,
                    table_ids,
                    hyperlink_ids,
                    relationship_roots,
                    relationship_records,
                )
                materialized_cell_count += comment_count
                _validate_materialized_cell_count(
                    sheet_name, materialized_cell_count
                )
                total_cells += cell_count
                if total_cells > MAX_PREFLIGHT_TOTAL_CELLS:
                    raise WorkbookStructureError(
                        "Workbook exceeds workbook cell-record limit of "
                        f"{MAX_PREFLIGHT_TOTAL_CELLS}"
                    )
                total_materialized_cells += materialized_cell_count
                if total_materialized_cells > MAX_PREFLIGHT_TOTAL_CELLS:
                    raise WorkbookStructureError(
                        "Workbook exceeds workbook cell materialization limit of "
                        f"{MAX_PREFLIGHT_TOTAL_CELLS}"
                    )
            _add_workbook_pivot_cache_roots(
                archive, relationship_roots, relationship_records
            )
            _preflight_relationship_graph(
                archive, relationship_roots, content_types, relationship_records
            )
    except WorkbookStructureError:
        raise
    except (BadZipFile, LargeZipFile, OSError, ValueError) as error:
        raise WorkbookStructureError(
            f"Workbook resource preflight failed: {error}"
        ) from error


def _read_workbook_sheet_refs(archive: ZipFile) -> list[tuple[str, str]]:
    refs: list[tuple[str, str]] = []
    try:
        with archive.open("xl/workbook.xml") as source:
            for _, element in iterparse(source, events=("end",)):
                if _xml_local_name(element.tag) == "sheet":
                    relationship_id = element.attrib.get(_DOCUMENT_RELATIONSHIP_ID)
                    if not relationship_id:
                        raise WorkbookStructureError(
                            "Workbook sheet is missing a relationship id"
                        )
                    refs.append(
                        (element.attrib.get("name") or "<unnamed>", relationship_id)
                    )
                    if len(refs) > MAX_PREFLIGHT_WORKSHEETS:
                        raise WorkbookStructureError(
                            "Workbook exceeds worksheet count limit of "
                            f"{MAX_PREFLIGHT_WORKSHEETS}"
                        )
                element.clear()
    except KeyError as error:
        raise WorkbookStructureError("Workbook XML is missing") from error
    except ParseError as error:
        raise WorkbookStructureError("Workbook XML is malformed") from error
    return refs


def _resolve_sheet_members(
    archive: ZipFile,
    sheet_refs: list[tuple[str, str]],
    content_types: _ContentTypes,
    relationship_records: _RelationshipRecordBudget,
) -> tuple[
    list[tuple[str, str, _Relationship]],
    list[tuple[str, str, _Relationship]],
]:
    wanted_ids = {relationship_id for _, relationship_id in sheet_refs}
    relationships: dict[str, tuple[str, str, str | None]] = {}
    part_records = 0
    count_toward_total = relationship_records.begins_part(
        "xl/_rels/workbook.xml.rels"
    )
    try:
        with archive.open("xl/_rels/workbook.xml.rels") as source:
            for _, element in iterparse(source, events=("end",)):
                if _xml_local_name(element.tag) == "Relationship":
                    part_records += 1
                    relationship_records.record(
                        "Workbook relationships", part_records, count_toward_total
                    )
                    relationship_id = element.attrib.get("Id")
                    if relationship_id in wanted_ids:
                        if relationship_id in relationships:
                            raise WorkbookStructureError(
                                "Workbook relationships contain a duplicate id"
                            )
                        relationships[relationship_id] = (
                            element.attrib.get("Type", ""),
                            element.attrib.get("Target", ""),
                            element.attrib.get("TargetMode"),
                        )
                element.clear()
    except KeyError as error:
        raise WorkbookStructureError("Workbook relationships XML is missing") from error
    except ParseError as error:
        raise WorkbookStructureError(
            "Workbook relationships XML is malformed"
        ) from error

    worksheets: list[tuple[str, str, _Relationship]] = []
    chartsheets: list[tuple[str, str, _Relationship]] = []
    for sheet_name, relationship_id in sheet_refs:
        relationship = relationships.get(relationship_id)
        if relationship is None:
            raise WorkbookStructureError(
                f"Workbook sheet {sheet_name} has no relationship"
            )
        relationship_type, target, target_mode = relationship
        relationship_object = _Relationship(
            relationship_id, relationship_type, target, target_mode
        )
        # OpenPyXL treats every workbook sheet relationship except a chartsheet
        # as a worksheet and eagerly parses its target.
        if "chartsheet" in relationship_type:
            if target_mode == "External":
                raise WorkbookStructureError(
                    f"Workbook sheet {sheet_name} has an external chartsheet relationship"
                )
            member_name = _safe_sheet_member(sheet_name, target, "chartsheets")
            _validate_relationship_target_semantics(
                content_types,
                _RelationshipTarget(
                    member_name,
                    relationship_id,
                    relationship_type,
                    "xl/workbook.xml",
                    "chartsheet",
                ),
            )
            chartsheets.append((sheet_name, member_name, relationship_object))
            continue
        if target_mode == "External":
            raise WorkbookStructureError(
                f"Workbook sheet {sheet_name} has an external worksheet relationship"
            )
        member_name = _safe_sheet_member(sheet_name, target, "worksheets")
        sheet_semantic = _relationship_semantic(relationship_type)
        expected_semantic = (
            sheet_semantic
            if sheet_semantic
            in {"worksheet", "dialogsheet", "macrosheet", "intlMacrosheet"}
            else "worksheet"
        )
        _validate_relationship_target_semantics(
            content_types,
            _RelationshipTarget(
                member_name,
                relationship_id,
                relationship_type,
                "xl/workbook.xml",
                expected_semantic,
            ),
        )
        worksheets.append((sheet_name, member_name, relationship_object))
    return worksheets, chartsheets


def _safe_sheet_member(sheet_name: str, target: str, sheet_kind: str) -> str:
    if (
        not target
        or "\\" in target
        or "\x00" in target
        or ":" in target
        or "?" in target
        or "#" in target
    ):
        raise WorkbookStructureError(
            f"Workbook sheet {sheet_name} has an unsafe {sheet_kind[:-1]} "
            "relationship target"
        )
    target_path = PurePosixPath(target)
    if ".." in target_path.parts:
        raise WorkbookStructureError(
            f"Workbook sheet {sheet_name} has an unsafe {sheet_kind[:-1]} "
            "relationship target"
        )
    candidate = target[1:] if target.startswith("/") else posix_join("xl", target)
    member_name = posix_normpath(candidate)
    parts = PurePosixPath(member_name).parts
    if len(parts) < 3 or parts[:2] != ("xl", sheet_kind):
        raise WorkbookStructureError(
            f"Workbook sheet {sheet_name} has an unsafe {sheet_kind[:-1]} "
            "relationship target"
        )
    return member_name


class _SizeLimitedReader:
    def __init__(self, source, limit: int, description: str) -> None:
        self._source = source
        self._limit = limit
        self._description = description
        self._bytes_read = 0

    def read(self, size: int = -1) -> bytes:
        chunk = self._source.read(size)
        self._bytes_read += len(chunk)
        if self._bytes_read > self._limit:
            raise WorkbookStructureError(
                f"{self._description} exceeds XML size limit of {self._limit}"
            )
        return chunk


def _preflight_chartsheet_xml(
    archive: ZipFile, sheet_name: str, member_name: str
) -> None:
    member = _required_xml_member(
        archive, member_name, f"Chartsheet XML is missing for {sheet_name}"
    )
    if member.file_size > MAX_PREFLIGHT_RELATED_XML_BYTES:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds chartsheet XML size limit of "
            f"{MAX_PREFLIGHT_RELATED_XML_BYTES}"
        )
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                f"{sheet_name} chartsheet",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Chartsheet XML is malformed for {sheet_name}"
        ) from error


def _preflight_worksheet_relationships(
    archive: ZipFile,
    sheet_name: str,
    worksheet_member: str,
    table_relationship_ids: set[str],
    hyperlink_relationship_ids: set[str],
    relationship_roots: dict[str, list[_RelationshipTarget]],
    relationship_records: _RelationshipRecordBudget,
) -> int:
    worksheet_path = PurePosixPath(worksheet_member)
    relationships_member = str(
        worksheet_path.parent
        / "_rels"
        / f"{worksheet_path.name}.rels"
    )
    try:
        member = archive.getinfo(relationships_member)
    except KeyError:
        if hyperlink_relationship_ids:
            raise WorkbookStructureError(
                f"Worksheet hyperlink relationships are missing for {sheet_name}"
            )
        if table_relationship_ids:
            raise WorkbookStructureError(
                f"Worksheet table relationships are missing for {sheet_name}"
            )
        return 0
    if member.is_dir():
        raise WorkbookStructureError(
            f"Worksheet relationships XML is invalid for {sheet_name}"
        )
    if member.file_size > MAX_PREFLIGHT_RELATED_XML_BYTES:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds worksheet relationships XML size limit of "
            f"{MAX_PREFLIGHT_RELATED_XML_BYTES}"
        )

    relationships: list[_Relationship] = []
    relationship_ids: set[str] = set()
    part_records = 0
    count_toward_total = relationship_records.begins_part(relationships_member)
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                f"{sheet_name} worksheet relationships",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "Relationship":
                    part_records += 1
                    relationship_records.record(
                        f"{sheet_name} worksheet relationships",
                        part_records,
                        count_toward_total,
                    )
                    relationship_id = element.attrib.get("Id")
                    if not relationship_id:
                        raise WorkbookStructureError(
                            f"Worksheet relationships are invalid for {sheet_name}"
                        )
                    relationship_ids.add(relationship_id)
                    relationship_type = element.attrib.get("Type", "")
                    relationships.append(
                        _Relationship(
                            relationship_id,
                            relationship_type,
                            element.attrib.get("Target", ""),
                            element.attrib.get("TargetMode"),
                        )
                    )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Worksheet relationships XML is malformed for {sheet_name}"
        ) from error
    if not table_relationship_ids.issubset(relationship_ids):
        raise WorkbookStructureError(
            f"Worksheet table relationships are missing for {sheet_name}"
        )
    if not hyperlink_relationship_ids.issubset(relationship_ids):
        raise WorkbookStructureError(
            f"Worksheet hyperlink relationships are missing for {sheet_name}"
        )

    relationships_by_id: dict[str, _Relationship] = {}
    for relationship in relationships:
        relationships_by_id.setdefault(relationship.id, relationship)
    for relationship_id in hyperlink_relationship_ids:
        _validate_hyperlink_relationship(
            relationships_by_id[relationship_id],
            f"{sheet_name} has an invalid worksheet hyperlink relationship",
        )

    eager_relationships: list[tuple[str, str]] = []
    for relationship in relationships:
        is_selected_table_relationship = (
            relationship.id in table_relationship_ids
            and relationships_by_id[relationship.id] is relationship
        )
        if relationship.target_mode == "External":
            if (
                relationship.type not in _HYPERLINK_RELATIONSHIPS
                or is_selected_table_relationship
            ):
                raise WorkbookStructureError(
                    f"{sheet_name} has an unsupported external "
                    "worksheet relationship"
                )
            _validate_hyperlink_relationship(
                relationship,
                f"{sheet_name} has an invalid worksheet hyperlink relationship",
            )
            continue
        target_member = _safe_worksheet_related_member(
            sheet_name,
            worksheet_member,
            relationship.target,
        )
        relationship_semantic = _relationship_semantic(relationship.type)
        if (
            relationship_semantic in _XML_RELATIONSHIP_CONTENT_TYPES
            or is_selected_table_relationship
        ):
            _add_relationship_root(
                relationship_roots,
                target_member,
                relationship,
                worksheet_member,
                "table" if is_selected_table_relationship else None,
            )
            eager_relationships.append((relationship.type, target_member))
            if len(eager_relationships) > MAX_PREFLIGHT_RELATIONSHIP_EDGES:
                raise WorkbookStructureError(
                    "Workbook exceeds relationship edge limit of "
                    f"{MAX_PREFLIGHT_RELATIONSHIP_EDGES}"
                )

    comment_count = 0
    for relationship_type, target_member in eager_relationships:
        if _relationship_semantic(relationship_type) == "comments":
            parts = PurePosixPath(target_member).parts
            if len(parts) < 3 or parts[:2] != ("xl", "comments"):
                raise WorkbookStructureError(
                    f"{sheet_name} has an unsafe comment relationship target"
                )
            related_member = _required_xml_member(
                archive,
                target_member,
                f"Comment XML is missing for {sheet_name}",
            )
            comment_count += _preflight_comment_xml(
                archive, sheet_name, related_member
            )
            continue
        _required_xml_member(
            archive,
            target_member,
            f"Worksheet related XML is missing for {sheet_name}",
        )
        # Tables, drawings, and pivot definitions do not create worksheet cells.
        # The relationship graph validates these eager parts and their descendants.
    return comment_count


def _safe_worksheet_related_member(
    sheet_name: str, worksheet_member: str, target: str
) -> str:
    if (
        not target
        or "\\" in target
        or "\x00" in target
        or ":" in target
        or "?" in target
        or "#" in target
    ):
        raise WorkbookStructureError(
            f"{sheet_name} has an unsafe worksheet relationship target"
        )
    parts = [] if target.startswith("/") else list(
        PurePosixPath(worksheet_member).parent.parts
    )
    for part in PurePosixPath(target).parts:
        if part in ("/", "."):
            continue
        if part == "..":
            if len(parts) <= 1:
                raise WorkbookStructureError(
                    f"{sheet_name} has an unsafe worksheet relationship target"
                )
            parts.pop()
            continue
        parts.append(part)
    if len(parts) < 2 or parts[0] != "xl":
        raise WorkbookStructureError(
            f"{sheet_name} has an unsafe worksheet relationship target"
        )
    return "/".join(parts)


def _validate_hyperlink_relationship(
    relationship: _Relationship, error_message: str
) -> None:
    if (
        relationship.type not in _HYPERLINK_RELATIONSHIPS
        or relationship.target_mode not in (None, "Internal", "External")
        or not relationship.target.strip()
        or len(relationship.target) > MAX_PREFLIGHT_HYPERLINK_TARGET_LENGTH
        or _HYPERLINK_CONTROL_CHARACTERS.search(relationship.target)
    ):
        raise WorkbookStructureError(error_message)


def _add_workbook_pivot_cache_roots(
    archive: ZipFile,
    relationship_roots: dict[str, list[_RelationshipTarget]],
    relationship_records: _RelationshipRecordBudget,
) -> None:
    cache_ids: list[str] = []
    try:
        with archive.open("xl/workbook.xml") as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                "Workbook pivot cache XML",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "pivotCache":
                    relationship_id = element.attrib.get(_DOCUMENT_RELATIONSHIP_ID)
                    if not relationship_id:
                        raise WorkbookStructureError(
                            "Workbook pivot cache is missing a relationship id"
                        )
                    cache_ids.append(relationship_id)
                    if len(cache_ids) > MAX_PREFLIGHT_RELATIONSHIP_QUEUE:
                        raise WorkbookStructureError(
                            "Workbook exceeds relationship queue limit of "
                            f"{MAX_PREFLIGHT_RELATIONSHIP_QUEUE}"
                        )
                element.clear()
    except KeyError as error:
        raise WorkbookStructureError("Workbook XML is missing") from error
    except ParseError as error:
        raise WorkbookStructureError("Workbook XML is malformed") from error

    wanted_ids = set(cache_ids)
    relationships: dict[str, _Relationship] = {}
    part_records = 0
    count_toward_total = relationship_records.begins_part(
        "xl/_rels/workbook.xml.rels"
    )
    try:
        with archive.open("xl/_rels/workbook.xml.rels") as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                "Workbook pivot cache relationships",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "Relationship":
                    part_records += 1
                    relationship_records.record(
                        "Workbook pivot cache relationships",
                        part_records,
                        count_toward_total,
                    )
                    relationship_id = element.attrib.get("Id")
                    if relationship_id in wanted_ids:
                        if relationship_id in relationships:
                            raise WorkbookStructureError(
                                "Workbook pivot cache relationships contain a duplicate id"
                            )
                        relationships[relationship_id] = _Relationship(
                            relationship_id,
                            element.attrib.get("Type", ""),
                            element.attrib.get("Target", ""),
                            element.attrib.get("TargetMode"),
                        )
                element.clear()
    except KeyError as error:
        raise WorkbookStructureError("Workbook relationships XML is missing") from error
    except ParseError as error:
        raise WorkbookStructureError(
            "Workbook relationships XML is malformed"
        ) from error

    for relationship_id in cache_ids:
        relationship = relationships.get(relationship_id)
        if relationship is None:
            raise WorkbookStructureError(
                "Workbook pivot cache has no relationship"
            )
        if relationship.target_mode == "External":
            raise WorkbookStructureError(
                "Workbook pivot cache has an unsupported external relationship"
            )
        _add_relationship_root(
            relationship_roots,
            _safe_relationship_graph_target(
                "xl/workbook.xml", relationship.target
            ),
            relationship,
            "xl/workbook.xml",
            "pivotCacheDefinition",
        )


def _add_relationship_root(
    relationship_roots: dict[str, list[_RelationshipTarget]],
    member_name: str,
    relationship: _Relationship,
    source_member: str,
    expected_semantic: str | None = None,
) -> None:
    relationship_target = _RelationshipTarget(
        member_name,
        relationship.id,
        relationship.type,
        source_member,
        expected_semantic,
    )
    existing_targets = relationship_roots.get(member_name)
    if existing_targets is not None:
        if relationship_target not in existing_targets:
            existing_targets.append(relationship_target)
        return
    if len(relationship_roots) >= MAX_PREFLIGHT_RELATIONSHIP_QUEUE:
        raise WorkbookStructureError(
            "Workbook exceeds relationship queue limit of "
            f"{MAX_PREFLIGHT_RELATIONSHIP_QUEUE}"
        )
    relationship_roots[member_name] = [relationship_target]


def _preflight_relationship_graph(
    archive: ZipFile,
    root_members: dict[str, list[_RelationshipTarget]],
    content_types: _ContentTypes,
    relationship_records: _RelationshipRecordBudget,
) -> None:
    if len(root_members) > MAX_PREFLIGHT_RELATIONSHIP_QUEUE:
        raise WorkbookStructureError(
            "Workbook exceeds relationship queue limit of "
            f"{MAX_PREFLIGHT_RELATIONSHIP_QUEUE}"
        )
    queue = deque(
        (relationship_target, 0)
        for targets in root_members.values()
        for relationship_target in targets
    )
    queued = {target.member_name for target, _ in queue}
    visited: set[str] = set()
    adjacency: dict[str, list[str]] = {}
    edge_count = 0

    while queue:
        relationship_target, depth = queue.popleft()
        member_name = relationship_target.member_name
        queued.discard(member_name)
        member = _required_graph_member(archive, member_name)
        is_xml = _validate_relationship_target_semantics(
            content_types, relationship_target
        )
        if member_name in visited:
            continue
        visited.add(member_name)
        if len(visited) > MAX_PREFLIGHT_RELATIONSHIP_PARTS:
            raise WorkbookStructureError(
                "Workbook exceeds relationship part limit of "
                f"{MAX_PREFLIGHT_RELATIONSHIP_PARTS}"
            )

        relationship_references: set[str] = set()
        hyperlink_references: set[str] = set()
        non_hyperlink_references: set[str] = set()
        expected_relationship_semantics: dict[str, str] = {}
        if is_xml:
            (
                relationship_references,
                hyperlink_references,
                non_hyperlink_references,
                expected_relationship_semantics,
            ) = _preflight_related_xml(
                archive,
                member_name,
                member,
            )
        elif member.file_size > MAX_PREFLIGHT_RELATED_PART_BYTES:
            raise WorkbookStructureError(
                f"{member_name} exceeds related part size limit of "
                f"{MAX_PREFLIGHT_RELATED_PART_BYTES}"
            )

        relationships, relationships_by_id = _read_graph_relationships(
            archive, member_name, relationship_records
        )
        if not relationship_references.issubset(relationships_by_id):
            raise WorkbookStructureError(
                f"{member_name} relationship references are missing"
            )
        for relationship_id in hyperlink_references:
            relationship = relationships_by_id[relationship_id]
            relationship_kind = (
                "drawing" if member_name.startswith("xl/drawings/") else "related XML"
            )
            _validate_hyperlink_relationship(
                relationship,
                f"{member_name} has an invalid {relationship_kind} "
                "hyperlink relationship",
            )
        for relationship_id in non_hyperlink_references:
            if relationships_by_id[relationship_id].target_mode == "External":
                raise WorkbookStructureError(
                    f"{member_name} has an unsupported external relationship"
                )
        targets = adjacency.setdefault(member_name, [])
        for relationship in relationships:
            edge_count += 1
            if edge_count > MAX_PREFLIGHT_RELATIONSHIP_EDGES:
                raise WorkbookStructureError(
                    "Workbook exceeds relationship edge limit of "
                    f"{MAX_PREFLIGHT_RELATIONSHIP_EDGES}"
                )
            if relationship.target_mode == "External":
                if relationship.type not in _HYPERLINK_RELATIONSHIPS:
                    raise WorkbookStructureError(
                        f"{member_name} has an unsupported external relationship"
                    )
                _validate_hyperlink_relationship(
                    relationship,
                    f"{member_name} has an invalid external hyperlink relationship",
                )
                continue
            target_member = _safe_relationship_graph_target(
                member_name, relationship.target
            )
            target_context = _RelationshipTarget(
                target_member,
                relationship.id,
                relationship.type,
                member_name,
                expected_relationship_semantics.get(relationship.id),
            )
            _required_graph_member(archive, target_member)
            _validate_relationship_target_semantics(
                content_types, target_context
            )
            targets.append(target_member)
            if target_member in visited or target_member in queued:
                continue
            target_depth = depth + 1
            if target_depth > MAX_PREFLIGHT_RELATIONSHIP_DEPTH:
                raise WorkbookStructureError(
                    "Workbook exceeds relationship depth limit of "
                    f"{MAX_PREFLIGHT_RELATIONSHIP_DEPTH}"
                )
            queue.append((target_context, target_depth))
            queued.add(target_member)
            if len(queue) > MAX_PREFLIGHT_RELATIONSHIP_QUEUE:
                raise WorkbookStructureError(
                    "Workbook exceeds relationship queue limit of "
                    f"{MAX_PREFLIGHT_RELATIONSHIP_QUEUE}"
                )

    _reject_relationship_cycles(adjacency)


def _read_graph_relationships(
    archive: ZipFile,
    source_member: str,
    relationship_records: _RelationshipRecordBudget,
) -> tuple[list[_Relationship], dict[str, _Relationship]]:
    source_path = PurePosixPath(source_member)
    relationships_member = str(
        source_path.parent / "_rels" / f"{source_path.name}.rels"
    )
    try:
        member = archive.getinfo(relationships_member)
    except KeyError:
        return [], {}
    if member.is_dir():
        raise WorkbookStructureError(
            f"Relationship graph relationships XML is invalid for {source_member}"
        )
    if member.file_size > MAX_PREFLIGHT_RELATED_XML_BYTES:
        raise WorkbookStructureError(
            f"{relationships_member} exceeds related XML size limit of "
            f"{MAX_PREFLIGHT_RELATED_XML_BYTES}"
        )

    relationships: list[_Relationship] = []
    part_records = 0
    count_toward_total = relationship_records.begins_part(relationships_member)
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                f"{source_member} relationship graph relationships",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "Relationship":
                    part_records += 1
                    relationship_records.record(
                        f"{source_member} relationship graph relationships",
                        part_records,
                        count_toward_total,
                    )
                    relationship_id = element.attrib.get("Id")
                    relationship_type = element.attrib.get("Type")
                    target = element.attrib.get("Target")
                    if not relationship_id or not relationship_type or not target:
                        raise WorkbookStructureError(
                            "Relationship graph relationships are invalid for "
                            f"{source_member}"
                        )
                    relationships.append(
                        _Relationship(
                            relationship_id,
                            relationship_type,
                            target,
                            element.attrib.get("TargetMode"),
                        )
                    )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            "Relationship graph relationships XML is malformed for "
            f"{source_member}"
        ) from error
    relationships_by_id: dict[str, _Relationship] = {}
    for relationship in relationships:
        relationships_by_id.setdefault(relationship.id, relationship)
    return relationships, relationships_by_id


def _safe_relationship_graph_target(source_member: str, target: str) -> str:
    if (
        not target
        or target.startswith("//")
        or "\\" in target
        or "\x00" in target
        or ":" in target
        or "?" in target
        or "#" in target
    ):
        raise WorkbookStructureError(
            f"{source_member} has an unsafe relationship graph target"
        )

    parts = [] if target.startswith("/") else list(
        PurePosixPath(source_member).parent.parts
    )
    for part in PurePosixPath(target).parts:
        if part in ("/", "."):
            continue
        if part == "..":
            if len(parts) <= 1:
                raise WorkbookStructureError(
                    f"{source_member} has an unsafe relationship graph target"
                )
            parts.pop()
            continue
        parts.append(part)
    if len(parts) < 2 or parts[0] != "xl":
        raise WorkbookStructureError(
            f"{source_member} has an unsafe relationship graph target"
        )
    return "/".join(parts)


def _required_graph_member(archive: ZipFile, member_name: str):
    try:
        member = archive.getinfo(member_name)
    except KeyError as error:
        raise WorkbookStructureError(
            f"Relationship graph target is missing: {member_name}"
        ) from error
    if member.is_dir():
        raise WorkbookStructureError(
            f"Relationship graph target is missing: {member_name}"
        )
    return member


def _relationship_semantic(relationship_type: str) -> str | None:
    alias = _RELATIONSHIP_SEMANTIC_ALIASES.get(relationship_type)
    if alias is not None:
        return alias
    for namespace in (
        _DOCUMENT_RELATIONSHIP_NAMESPACE,
        _STRICT_DOCUMENT_RELATIONSHIP_NAMESPACE,
    ):
        prefix = f"{namespace}/"
        if relationship_type.startswith(prefix):
            semantic = relationship_type[len(prefix) :]
            return semantic or None
    return None


def _is_xml_content_type(content_type: str) -> bool:
    lowered = content_type.lower()
    return (
        lowered in {"application/xml", "text/xml"}
        or lowered.endswith("+xml")
        or lowered == "application/vnd.openxmlformats-officedocument.vmldrawing"
    )


def _validate_relationship_target_semantics(
    content_types: _ContentTypes, target: _RelationshipTarget
) -> bool:
    content_type = content_types.for_member(target.member_name)
    relationship_semantic = _relationship_semantic(target.relationship_type)
    semantic = target.expected_semantic or relationship_semantic
    semantic_description = semantic or target.relationship_type
    if content_type is None:
        raise WorkbookStructureError(
            f"{target.member_name} content type mapping is missing for "
            f"{semantic_description} relationship {target.relationship_id}"
        )

    semantics = [semantic]
    if relationship_semantic not in semantics:
        semantics.append(relationship_semantic)
    result: bool | None = None
    for candidate_semantic in semantics:
        expected_xml_types = _XML_RELATIONSHIP_CONTENT_TYPES.get(
            candidate_semantic or ""
        )
        if expected_xml_types is not None:
            if content_type not in expected_xml_types:
                raise WorkbookStructureError(
                    f"{target.member_name} content type conflicts with "
                    f"{candidate_semantic} relationship semantics"
                )
            result = True
            continue

        expected_binary_prefixes = _BINARY_RELATIONSHIP_CONTENT_TYPE_PREFIXES.get(
            candidate_semantic or ""
        )
        if expected_binary_prefixes is not None:
            if not content_type.startswith(expected_binary_prefixes):
                raise WorkbookStructureError(
                    f"{target.member_name} content type conflicts with "
                    f"{candidate_semantic} relationship semantics"
                )
            if result is True:
                raise WorkbookStructureError(
                    f"{target.member_name} has conflicting relationship semantics"
                )
            result = False

    if result is not None:
        return result
    return _is_xml_content_type(content_type)


def _preflight_related_xml(
    archive: ZipFile, member_name: str, member
) -> tuple[set[str], set[str], set[str], dict[str, str]]:
    if member.file_size > MAX_PREFLIGHT_RELATED_XML_BYTES:
        raise WorkbookStructureError(
            f"{member_name} exceeds related XML size limit of "
            f"{MAX_PREFLIGHT_RELATED_XML_BYTES}"
        )
    relationship_references: set[str] = set()
    hyperlink_references: set[str] = set()
    non_hyperlink_references: set[str] = set()
    expected_relationship_semantics: dict[str, str] = {}
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                f"{member_name} related XML",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                local_name = _xml_local_name(element.tag)
                for attribute_name in _DOCUMENT_RELATIONSHIP_REFERENCES:
                    relationship_id = element.attrib.get(attribute_name)
                    if relationship_id:
                        relationship_references.add(relationship_id)
                        if local_name in (
                            "hlinkClick",
                            "hlinkHover",
                            "hlinkMouseOver",
                        ):
                            hyperlink_references.add(relationship_id)
                        else:
                            non_hyperlink_references.add(relationship_id)
                        expected_semantic = _expected_relationship_semantic(
                            local_name, attribute_name
                        )
                        previous_semantic = expected_relationship_semantics.get(
                            relationship_id
                        )
                        if (
                            expected_semantic is not None
                            and previous_semantic is not None
                            and previous_semantic != expected_semantic
                        ):
                            raise WorkbookStructureError(
                                f"{member_name} relationship {relationship_id} has "
                                "conflicting consumer semantics"
                            )
                        if expected_semantic is not None:
                            expected_relationship_semantics[relationship_id] = (
                                expected_semantic
                            )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Related XML is malformed: {member_name}"
        ) from error
    return (
        relationship_references,
        hyperlink_references,
        non_hyperlink_references,
        expected_relationship_semantics,
    )


def _expected_relationship_semantic(
    local_name: str, attribute_name: str
) -> str | None:
    if local_name in ("hlinkClick", "hlinkHover", "hlinkMouseOver"):
        return "hyperlink"
    if local_name == "blip" and attribute_name.endswith("}embed"):
        return "image"
    if local_name == "chart" and attribute_name.endswith("}id"):
        return "chart"
    if local_name == "pivotCacheDefinition" and attribute_name.endswith("}id"):
        return "pivotCacheRecords"
    return None


def _reject_relationship_cycles(adjacency: dict[str, list[str]]) -> None:
    states: dict[str, int] = {}
    for root in adjacency:
        if states.get(root, 0) != 0:
            continue
        states[root] = 1
        stack = [(root, iter(adjacency.get(root, ())))]
        while stack:
            _, targets = stack[-1]
            try:
                target = next(targets)
            except StopIteration:
                completed, _ = stack.pop()
                states[completed] = 2
                continue
            state = states.get(target, 0)
            if state == 1:
                raise WorkbookStructureError(
                    "Workbook relationship graph contains a cycle"
                )
            if state == 2:
                continue
            states[target] = 1
            stack.append((target, iter(adjacency.get(target, ()))))


def _required_xml_member(archive: ZipFile, member_name: str, message: str):
    try:
        member = archive.getinfo(member_name)
    except KeyError as error:
        raise WorkbookStructureError(message) from error
    if member.is_dir():
        raise WorkbookStructureError(message)
    return member


def _preflight_comment_xml(archive: ZipFile, sheet_name: str, member) -> int:
    if member.file_size > MAX_PREFLIGHT_RELATED_XML_BYTES:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds comment XML size limit of "
            f"{MAX_PREFLIGHT_RELATED_XML_BYTES}"
        )
    comment_count = 0
    try:
        with archive.open(member) as source:
            limited_source = _SizeLimitedReader(
                source,
                MAX_PREFLIGHT_RELATED_XML_BYTES,
                f"{sheet_name} comment",
            )
            for _, element in iterparse(limited_source, events=("end",)):
                if _xml_local_name(element.tag) == "comment":
                    comment_count += 1
                    _validate_materialized_cell_count(sheet_name, comment_count)
                    _validate_preflight_cell_reference(
                        element.attrib.get("ref"), sheet_name
                    )
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Comment XML is malformed for {sheet_name}"
        ) from error
    return comment_count


def _preflight_worksheet_xml(
    archive: ZipFile, sheet_name: str, member_name: str
) -> tuple[int, int, set[str], set[str]]:
    try:
        member = archive.getinfo(member_name)
    except KeyError as error:
        raise WorkbookStructureError(
            f"Worksheet XML is missing for {sheet_name}"
        ) from error
    if member.is_dir():
        raise WorkbookStructureError(f"Worksheet XML is invalid for {sheet_name}")
    if member.file_size > MAX_PREFLIGHT_WORKSHEET_XML_BYTES:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds worksheet XML size limit of "
            f"{MAX_PREFLIGHT_WORKSHEET_XML_BYTES}"
        )

    cell_count = 0
    materialized_cell_count = 0
    row_record_count = 0
    column_dimension_count = 0
    inferred_row_number = 0
    table_relationship_ids: set[str] = set()
    hyperlink_relationship_ids: set[str] = set()
    try:
        with archive.open(member) as source:
            for _, element in iterparse(source, events=("end",)):
                local_name = _xml_local_name(element.tag)
                if local_name == "c":
                    cell_count += 1
                    if cell_count > MAX_PREFLIGHT_WORKSHEET_CELLS:
                        raise WorkbookStructureError(
                            f"{sheet_name} exceeds worksheet cell-record limit of "
                            f"{MAX_PREFLIGHT_WORKSHEET_CELLS}"
                        )
                    materialized_cell_count += 1
                    _validate_materialized_cell_count(
                        sheet_name, materialized_cell_count
                    )
                    _validate_preflight_cell_reference(
                        element.attrib.get("r"), sheet_name
                    )
                elif local_name == "dimension":
                    reference = element.attrib.get("ref")
                    if reference:
                        _preflight_range_area(reference, sheet_name)
                elif local_name == "mergeCell":
                    materialized_cell_count += _preflight_range_area(
                        element.attrib.get("ref"), sheet_name
                    )
                    _validate_materialized_cell_count(
                        sheet_name, materialized_cell_count
                    )
                elif local_name == "hyperlink":
                    materialized_cell_count += _preflight_range_area(
                        element.attrib.get("ref"), sheet_name
                    )
                    _validate_materialized_cell_count(
                        sheet_name, materialized_cell_count
                    )
                    relationship_id = element.attrib.get(
                        _DOCUMENT_RELATIONSHIP_ID
                    )
                    if relationship_id:
                        hyperlink_relationship_ids.add(relationship_id)
                elif local_name == "row":
                    row_record_count += 1
                    if row_record_count > MAX_PREFLIGHT_ROWS:
                        raise WorkbookStructureError(
                            f"{sheet_name} exceeds preflight row limit of "
                            f"{MAX_PREFLIGHT_ROWS} (worksheet row-record limit of "
                            f"{MAX_PREFLIGHT_ROWS})"
                        )
                    row_reference = element.attrib.get("r")
                    if row_reference is None:
                        row_number = inferred_row_number + 1
                    else:
                        try:
                            row_number = int(row_reference)
                        except ValueError as error:
                            raise WorkbookStructureError(
                                f"{sheet_name} contains an invalid row reference"
                            ) from error
                    inferred_row_number = row_number
                    if not 1 <= row_number <= MAX_PREFLIGHT_ROWS:
                        raise WorkbookStructureError(
                            f"{sheet_name} exceeds preflight row limit of "
                            f"{MAX_PREFLIGHT_ROWS}"
                        )
                elif local_name == "col":
                    column_dimension_count += 1
                    if (
                        column_dimension_count
                        > MAX_PREFLIGHT_COLUMN_DIMENSION_RECORDS
                    ):
                        raise WorkbookStructureError(
                            f"{sheet_name} exceeds worksheet column-dimension "
                            "record limit of "
                            f"{MAX_PREFLIGHT_COLUMN_DIMENSION_RECORDS}"
                        )
                    _validate_preflight_column_dimension(element.attrib, sheet_name)
                elif local_name == "tablePart":
                    relationship_id = element.attrib.get(_DOCUMENT_RELATIONSHIP_ID)
                    if not relationship_id:
                        raise WorkbookStructureError(
                            f"Worksheet table relationship is invalid for {sheet_name}"
                        )
                    table_relationship_ids.add(relationship_id)
                element.clear()
    except ParseError as error:
        raise WorkbookStructureError(
            f"Worksheet XML is malformed for {sheet_name}"
        ) from error
    return (
        cell_count,
        materialized_cell_count,
        table_relationship_ids,
        hyperlink_relationship_ids,
    )


def _validate_preflight_column_dimension(
    attributes: dict[str, str], sheet_name: str
) -> None:
    minimum = _parse_bounded_decimal(attributes.get("min"), MAX_EXCEL_COLUMNS)
    maximum = _parse_bounded_decimal(attributes.get("max"), MAX_EXCEL_COLUMNS)
    if minimum is None or maximum is None or maximum < minimum:
        raise WorkbookStructureError(
            f"{sheet_name} contains an invalid column dimension"
        )


def _parse_bounded_decimal(value: str | None, maximum: int) -> int | None:
    if (
        value is None
        or not value
        or len(value) > len(str(maximum))
        or not value.isascii()
        or not value.isdecimal()
    ):
        return None
    parsed = int(value)
    return parsed if 1 <= parsed <= maximum else None


def _validate_materialized_cell_count(sheet_name: str, count: int) -> None:
    if count > MAX_PREFLIGHT_WORKSHEET_CELLS:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds worksheet cell materialization limit of "
            f"{MAX_PREFLIGHT_WORKSHEET_CELLS}"
        )


def _preflight_range_area(reference: str | None, sheet_name: str) -> int:
    coordinates = (reference or "").split(":")
    if len(coordinates) not in (1, 2):
        raise WorkbookStructureError(
            f"{sheet_name} contains an invalid cell range reference"
        )
    start_row, start_column = _validate_preflight_cell_reference(
        coordinates[0], sheet_name
    )
    if len(coordinates) == 1:
        return 1
    end_row, end_column = _validate_preflight_cell_reference(
        coordinates[1], sheet_name
    )
    if end_row < start_row or end_column < start_column:
        raise WorkbookStructureError(
            f"{sheet_name} contains an invalid cell range reference"
        )
    return (end_row - start_row + 1) * (end_column - start_column + 1)


def _validate_preflight_cell_reference(
    reference: str | None, sheet_name: str
) -> tuple[int, int]:
    match = _CELL_REFERENCE.fullmatch(reference or "")
    if match is None:
        raise WorkbookStructureError(
            f"{sheet_name} contains an invalid cell reference"
        )
    row_number = int(match.group("row"))
    column_number = 0
    for character in match.group("column").upper():
        column_number = column_number * 26 + ord(character) - ord("A") + 1
        if column_number > MAX_PREFLIGHT_COLUMNS:
            break
    if row_number > MAX_PREFLIGHT_ROWS:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds preflight row limit of {MAX_PREFLIGHT_ROWS}"
        )
    if column_number > MAX_PREFLIGHT_COLUMNS:
        raise WorkbookStructureError(
            f"{sheet_name} exceeds preflight column limit of {MAX_PREFLIGHT_COLUMNS}"
        )
    return row_number, column_number


def _xml_local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _append_parsed_entry(
    rows: list[ParsedTransportEntry], entry: ParsedTransportEntry, limit: int
) -> None:
    if len(rows) >= limit:
        raise WorkbookStructureError(
            f"Workbook parsed entry limit of {MAX_PARSED_ENTRIES} exceeded"
        )
    rows.append(entry)


def _find_regular_header(sheet, formula_sheet) -> int | None:
    for row_number in range(1, min(sheet.max_row, 50) + 1):
        if _is_regular_header_view(sheet, row_number):
            if not _is_regular_header(sheet, formula_sheet, row_number):
                return None
            return row_number
    return None


def _is_regular_header(sheet, formula_sheet, row_number: int) -> bool:
    if not all(
        _is_regular_header_view(workbook_sheet, row_number)
        for workbook_sheet in (sheet, formula_sheet)
    ):
        return False
    if any(
        _has_value(workbook_sheet.cell(row_number, 35).value)
        for workbook_sheet in (sheet, formula_sheet)
    ):
        return False
    return all(
        sheet.cell(row_number, column).value
        == formula_sheet.cell(row_number, column).value
        and not _is_formula(formula_sheet.cell(row_number, column).value)
        for column in (1, 38)
    )


def _is_regular_header_view(sheet, row_number: int) -> bool:
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
    sheet, formula_sheet
) -> tuple[int, int, int, int, int | None, int] | None:
    for row_number in range(1, min(sheet.max_row, 50) + 1):
        header = _subcontract_header_view(sheet, row_number)
        if header is None:
            continue
        if (
            _subcontract_header_view(formula_sheet, row_number) != header
            or any(
                _is_formula(formula_sheet.cell(row_number, column).value)
                for column in range(1, _SUBCONTRACT_LAST_FOOTPRINT_COLUMN + 1)
            )
        ):
            return None
        return (row_number, *header, _SUBCONTRACT_LAST_FOOTPRINT_COLUMN)
    return None


def _subcontract_header_view(
    sheet, row_number: int
) -> tuple[int, int, int, int | None] | None:
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
            columns["date"],
            columns["destination"],
            columns["tonnage"],
            columns.get("number"),
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


def _is_displayed_total_row(sheet, formula_sheet, row_number: int) -> bool:
    label_columns = [
        column
        for column in _SUBCONTRACT_TOTAL_LABEL_COLUMNS
        if _header_text(sheet.cell(row_number, column).value) in _TOTAL_LABELS
    ]
    if len(label_columns) != 1:
        return False
    allowed_columns = {label_columns[0], 9}
    return any(
        _has_value(workbook_sheet.cell(row_number, 9).value)
        for workbook_sheet in (sheet, formula_sheet)
    ) and not any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in range(1, _SUBCONTRACT_LAST_FOOTPRINT_COLUMN + 1)
        if column not in allowed_columns
    )


def _is_regular_footer_or_note(sheet, formula_sheet, row_number: int) -> bool:
    if not _has_regular_footer_label(sheet, formula_sheet, row_number):
        return False
    cached_marker = sheet.cell(row_number, 38).value
    formula_marker = formula_sheet.cell(row_number, 38).value
    return (
        not isinstance(cached_marker, bool)
        and isinstance(cached_marker, (int, float, Decimal))
        and (
            _is_formula(formula_marker)
            or (
                not isinstance(formula_marker, bool)
                and isinstance(formula_marker, (int, float, Decimal))
            )
        )
    )


def _has_regular_footer_label(sheet, formula_sheet, row_number: int) -> bool:
    if any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in range(1, 37)
    ):
        return False
    cached_label = sheet.cell(row_number, 37).value
    formula_label = formula_sheet.cell(row_number, 37).value
    if not all(
        isinstance(value, str) and bool(value.strip()) and not _is_formula(value)
        for value in (cached_label, formula_label)
    ):
        return False
    return True


def _is_regular_blank_or_group_row(sheet, formula_sheet, row_number: int) -> bool:
    cached_group = sheet.cell(row_number, 1).value
    formula_group = formula_sheet.cell(row_number, 1).value
    if _has_value(formula_group) and not _has_value(cached_group):
        return False
    return not any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in range(2, 39)
    )


def _is_regular_inactive_row(sheet, formula_sheet, row_number: int) -> bool:
    cached_group = sheet.cell(row_number, 1).value
    formula_group = formula_sheet.cell(row_number, 1).value
    if _has_value(formula_group) and not _has_value(cached_group):
        return False
    if any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in (35, 38)
    ):
        return False
    return (
        not _has_value(sheet.cell(row_number, 36).value)
        and all(
            _is_blank_or_numeric_zero(sheet.cell(row_number, column).value)
            for column in range(3, 34)
        )
        and _is_blank_or_numeric_zero(sheet.cell(row_number, 34).value)
        and _is_blank_or_numeric_zero(sheet.cell(row_number, 37).value)
    )


def _is_subcontract_template_row(
    sheet,
    formula_sheet,
    row_number: int,
    number_column: int | None,
    last_footprint_column: int,
) -> bool:
    if number_column is None:
        return False
    return _has_value(sheet.cell(row_number, number_column).value) and not any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in range(1, last_footprint_column + 1)
        if column != number_column
    )


def _is_subcontract_formula_footer(
    sheet, formula_sheet, row_number: int, last_footprint_column: int
) -> bool:
    return _is_formula(formula_sheet.cell(row_number, 9).value) and not any(
        _has_value(workbook_sheet.cell(row_number, column).value)
        for workbook_sheet in (sheet, formula_sheet)
        for column in range(1, last_footprint_column + 1)
        if column != 9
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
