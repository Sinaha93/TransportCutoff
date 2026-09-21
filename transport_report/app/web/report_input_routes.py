"""Explicit current-report supplements and next-month plan entry."""
from dataclasses import replace

from fastapi import APIRouter, Request
from fastapi.responses import RedirectResponse

from app.repositories.report_inputs import (
    ReportInputRepository, ReportSupplementalInputs, next_report_month,
)
from app.web.routes import FormError, database, form, integer, quantity, render, require_month


router = APIRouter()


@router.get("/months/{month}/report-inputs")
def report_inputs(request: Request, month: str):
    require_month(month)
    snapshot = ReportInputRepository(database(request)).get_form_snapshot(month)
    return render(request, "report_inputs.html", snapshot=snapshot,
        plans={row["destination_id"]: row for row in snapshot.plans})


def _inputs(data, prefix, month):
    def one(name):
        values = data.getlist(f"{prefix}_{name}")
        if len(values) != 1:
            raise FormError("보고 보조 입력 항목이 누락되거나 중복되었습니다. 화면을 다시 여세요.")
        return values[0]

    actual = prefix == "current"
    return ReportSupplementalInputs(
        month, planned_sales_won=integer(one("planned_sales"), "계획 매출", blank=True),
        nonregular_planned_quantity_ea=quantity(one("nonregular_planned_quantity"), "비정규 계획 수량"),
        nonregular_planned_cost_won=integer(one("nonregular_planned_cost"), "비정규 계획 비용", blank=True),
        nonregular_actual_quantity_ea=quantity(one("nonregular_actual_quantity"), "비정규 실적 수량") if actual else None,
        nonregular_actual_cost_won=integer(one("nonregular_actual_cost"), "비정규 실적 비용", blank=True) if actual else None,
        source_note=one("source_note"),
    )


@router.post("/months/{month}/report-inputs")
async def save_report_inputs(request: Request, month: str):
    require_month(month)
    data = await form(request, "save_report_inputs")
    if len(data.getlist("supplemental_revision")) != 1 or not data.get("supplemental_revision"):
        raise FormError("보고 보조 입력 검토 정보가 없습니다. 화면을 다시 여세요.", 409)
    current = _inputs(data, "current", month)
    next_inputs = _inputs(data, "next", next_report_month(month))
    keys = ("destination_id", "plan_quantity", "plan_cost", "representative_item")
    columns = [data.getlist(key) for key in keys]
    if len({len(column) for column in columns}) != 1:
        raise FormError("납품처별 계획 입력 행이 누락되었습니다. 화면을 다시 여세요.")
    rows = []
    for values in zip(*columns):
        row = dict(zip(keys, values))
        row.update(destination_id=integer(row["destination_id"], "납품처"),
            plan_quantity=quantity(row["plan_quantity"], "다음 달 계획 수량"),
            plan_cost=integer(row["plan_cost"], "다음 달 계획 비용", blank=True))
        rows.append(row)
    repo = ReportInputRepository(database(request))
    # This screen edits next-month plans only. The revision also protects these
    # carried-forward actuals from a concurrent change before the atomic write.
    stored_next = repo.get(next_inputs.report_month)
    if stored_next is not None:
        next_inputs = replace(next_inputs,
            nonregular_actual_quantity_ea=stored_next.nonregular_actual_quantity_ea,
            nonregular_actual_cost_won=stored_next.nonregular_actual_cost_won)
    repo.save_form(month, current, next_inputs, rows, expected_revision=data["supplemental_revision"])
    return RedirectResponse(f"/months/{month}/report-inputs", status_code=303)
