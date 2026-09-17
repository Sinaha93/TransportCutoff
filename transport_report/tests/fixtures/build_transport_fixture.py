from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook


REGULAR_SHEET = "화성운반비내역(8월)"
SUBCONTRACT_SHEETS = (
    "용차1-OK로지웰",
    "용차2-대원로지스틱",
    "용차3-정동물류",
)


def build_transport_fixture(
    path: Path, *, subtotal_mismatch: bool = False
) -> Path:
    workbook = Workbook()
    regular = workbook.active
    regular.title = REGULAR_SHEET
    _write_regular_header(regular, 2)

    regular.cell(3, 1, "5톤 / 기사A")
    regular.cell(3, 2, "Known Plant")
    regular.cell(3, 14, 1.25)  # day 12 (C is day 1)
    regular.cell(3, 34, 1.25)  # AH: total count
    regular.cell(3, 36, 2_000)  # AJ: unit cost
    regular.cell(3, 37, 2_501 if subtotal_mismatch else 2_500)  # AK: subtotal

    _write_regular_header(regular, 4)
    regular.row_dimensions[4].hidden = True

    regular.cell(5, 2, "Unknown Plant")
    regular.cell(5, 3, 1)
    regular.cell(5, 34, 1)
    regular.cell(5, 36, 3_000)
    regular.cell(5, 37, 3_000)

    for offset, sheet_name in enumerate(SUBCONTRACT_SHEETS, start=5):
        sheet = workbook.create_sheet(sheet_name)
        sheet.cell(1, 1, "일자")
        sheet.cell(1, 2, "운반지역")
        sheet.cell(1, 8, "톤수")
        sheet.cell(1, 9, "금액")
        sheet.cell(2, 1, date(2026, 8, offset))
        sheet.cell(2, 2, "Known Plant")
        sheet.cell(2, 8, f"{offset}톤")
        sheet.cell(2, 9, offset * 10_000)
        sheet.cell(3, 2, "합계")
        sheet.cell(3, 9, offset * 10_000)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()
    return path


def _write_regular_header(sheet, row: int) -> None:
    sheet.cell(row, 1, "차량/기사")
    sheet.cell(row, 2, "운반지역")
    for day in range(1, 32):
        sheet.cell(row, day + 2, day)
    sheet.cell(row, 34, "총회수")
    sheet.cell(row, 36, "단가")
    sheet.cell(row, 37, "소계")
