"""Local HTML forms and server-validated report generation."""
from __future__ import annotations

import re
import secrets
import sqlite3
from decimal import Decimal
from pathlib import Path
from urllib.parse import parse_qs

from fastapi import APIRouter, FastAPI, Request
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from starlette.datastructures import FormData

from app.db import Database
from app.config import RuntimePaths
from app.services.report_service import ReportGenerationError, ReportService
from app.domain.calculations import (
    MAX_QUANTITY_EA, MAX_QUANTITY_SCALE, MAX_SQLITE_INTEGER,
    HistoricalValueKind, calculate_destination, calculate_group,
    calculate_total, historical_averages,
)
from app.domain.validation import (
    DestinationValidationInput, ImportBatchInput,
    ReportMonthSources, SalesInput, UnresolvedDestinationAlias,
    ValidationContext, ValidationIssue, ValidationResult, validate_report,
)
from app.importers.legacy_migration import (
    LegacyMigrationError, LegacyMigrationRevisionError, LegacyMigrationService,
)
from app.repositories.masters import MasterDataError, MasterRepository
from app.repositories.monthly_inputs import (
    MonthlyInputError, MonthlyInputRepository, MonthlyInputRevisionError, MonthlySales,
)
from app.repositories.transport_entries import TransportEntryRepository


MONTH = "2026-08"
templates = Jinja2Templates(directory=Path(__file__).parent / "templates")
templates.env.filters["value"] = lambda value: "미입력" if value is None else str(value)
router = APIRouter()


class FormError(ValueError):
    def __init__(self, message: str, status: int = 422):
        super().__init__(message)
        self.status = status


def database(request: Request) -> Database:
    if not request.app.state.database_ready:
        request.app.state.database.migrate()
        try:
            ReportService(request.app.state.database, request.app.state.runtime_paths).recover_pending()
        except ReportGenerationError as error:
            raise FormError(str(error), 409) from error
        request.app.state.database_ready = True
    return request.app.state.database


def require_month(month: str) -> None:
    if month != MONTH:
        raise FormError("현재 입력·보고 범위는 2026-08입니다.")


def render(request: Request, template: str, *, status: int = 200, **context):
    token = request.cookies.get("csrf_token") or secrets.token_urlsafe(32)
    response = templates.TemplateResponse(
        request=request, name=template,
        context={"month": MONTH, "csrf_token": token, **context}, status_code=status,
    )
    response.set_cookie("csrf_token", token, httponly=True, samesite="strict")
    response.headers["Cache-Control"] = "no-store"
    return response


async def form(request: Request, intent: str) -> FormData:
    # Native URL-encoded forms avoid adding a multipart dependency for local paths.
    if request.headers.get("content-type", "").split(";", 1)[0] != "application/x-www-form-urlencoded":
        raise FormError("입력 양식 형식이 올바르지 않습니다.")
    values = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    data = FormData((key, value) for key, items in values.items() for value in items)
    token = data.get("csrf_token", "")
    if len(data.getlist("csrf_token")) != 1 or not token or not secrets.compare_digest(token, request.cookies.get("csrf_token", "")):
        raise FormError("화면 확인 정보가 만료되었습니다. 화면을 다시 열고 저장하세요.", 403)
    if data.getlist("intent") != [intent]:
        raise FormError("입력 작업 구분이 올바르지 않습니다.")
    require_month(data.get("month", MONTH))
    return data


def integer(value: str, label: str, *, blank: bool = False) -> int | None:
    if blank and not value.strip():
        return None
    if not re.fullmatch(r"[0-9]+", value.strip()) or len(value.strip()) > 19:
        raise FormError(f"{label}: 0 이상의 정수를 입력하세요.")
    result = int(value)
    if result > MAX_SQLITE_INTEGER:
        raise FormError(f"{label}: 입력 가능 범위를 초과했습니다.")
    return result


def quantity(value: str, label: str) -> Decimal | None:
    text = value.strip()
    if not text:
        return None
    if not re.fullmatch(r"[0-9]+(?:\.[0-9]{1,4})?", text) or len(text) > 20:
        raise FormError(f"{label}: 0 이상, 소수 {MAX_QUANTITY_SCALE}자리 이내 수량을 입력하세요.")
    result = Decimal(text)
    if result > MAX_QUANTITY_EA:
        raise FormError(f"{label}: 수량 입력 한도를 초과했습니다.")
    return result


def preview_redirect():
    return RedirectResponse(f"/months/{MONTH}/preview", status_code=303)


def report_preview_data(db: Database, month: str) -> dict:
    """Adapt persisted values to the established pure calculators and validator.

    No export metadata is fabricated. Missing pairs stay missing in calculations;
    the raw values and blocking issues remain visible for correction.
    """
    require_month(month)
    masters = MasterRepository(db)
    inputs = MonthlyInputRepository(db)
    destinations = masters.list_destinations()
    active = [destination for destination in destinations if destination.active]
    plans = {item.destination_id: item for item in inputs.list_plans(month)}
    actuals = {item.destination_id: item for item in inputs.list_actual_quantities(month)}
    sales = inputs.get_sales(month)
    batches = TransportEntryRepository(db).list_current_batches(month)
    with db.connection() as connection:
        costs = {row["destination_id"]: row["cost_won"] for row in connection.execute("SELECT * FROM monthly_actual_costs WHERE report_month = ?", (month,))}
        # Only selected current batches override migrated monthly costs.
        for row in connection.execute(
            "SELECT t.destination_id, SUM(t.cost_won) AS cost_won FROM transport_entries t "
            "JOIN import_batches b ON b.id = t.import_batch_id WHERE t.report_month = ? AND b.is_current = 1 AND t.destination_id IS NOT NULL GROUP BY t.destination_id", (month,),
        ):
            costs[row["destination_id"]] = row["cost_won"]
        unknown = tuple(UnresolvedDestinationAlias(row["unresolved_alias"], f'{row["source_filename"]}: {row["source_sheet"]}!행{row["source_row"]}') for row in connection.execute(
            "SELECT t.*, b.source_filename FROM transport_entries t JOIN import_batches b ON b.id = t.import_batch_id WHERE t.report_month = ? AND b.is_current = 1 AND t.destination_id IS NULL", (month,),
        ))
        history_quantities = {(row["report_month"], row["destination_id"]): Decimal(row["quantity_ea_text"]) for row in connection.execute("SELECT * FROM monthly_actual_quantities WHERE report_month < ?", (month,))}
        history_costs = {(row["report_month"], row["destination_id"]): row["cost_won"] for row in connection.execute("SELECT * FROM monthly_actual_costs WHERE report_month < ?", (month,))}
        for row in connection.execute(
            "SELECT t.report_month, t.destination_id, SUM(t.cost_won) AS cost_won FROM transport_entries t JOIN import_batches b ON b.id = t.import_batch_id WHERE t.report_month < ? AND b.is_current = 1 AND t.destination_id IS NOT NULL GROUP BY t.report_month, t.destination_id", (month,),
        ):
            history_costs[row["report_month"], row["destination_id"]] = row["cost_won"]
    historical = {}
    for label, values, flag, kind in (
        ("quantity", history_quantities, "include_quantity_total", HistoricalValueKind.QUANTITY),
        ("cost", history_costs, "include_cost_total", HistoricalValueKind.MONEY),
    ):
        monthly = {}
        for historical_month in {key[0] for key in values}:
            selected = [values.get((historical_month, destination.id)) for destination in active if getattr(destination, flag)]
            monthly[historical_month] = (sum(selected, Decimal(0)) if selected and all(value is not None for value in selected) else None)
        historical[label] = historical_averages(month, monthly, value_kind=kind)
    calculations = {}
    issues = []
    rows = []
    for destination in active:
        plan = plans.get(destination.id)
        actual = actuals.get(destination.id)
        cost = costs.get(destination.id)
        actual_quantity = actual.quantity_ea if actual else None
        result = calculate_destination(
            plan.quantity_ea if plan else None, plan.cost_won if plan else None,
            actual_quantity if cost is not None else None,
            cost if actual_quantity is not None else None, destination_id=destination.id,
        )
        calculations[destination.id] = result
        rows.append({"destination": destination, "result": result, "actual_quantity": actual_quantity, "actual_cost": cost})
        for missing, code, label in ((plan is None, "MISSING_CURRENT_PLAN", "당월 계획"), (cost is None, "MISSING_ACTUAL_COST", "당월 운반비")):
            if destination.required_for_report and missing:
                issues.append(ValidationIssue(code, "error", f"{destination.name}: {label} 미입력입니다.", destination.id))
    context = ValidationContext(
        report_month=month,
        sales=SalesInput(sales.amount_won, sales.confirmed_at, "월 매출 입력") if sales else None,
        prior_year_history=historical["quantity"].comparison_year,
        month_sources=ReportMonthSources(month if plans else None, month if actuals else None, month if batches else None, None, None),
        unresolved_aliases=unknown,
        # Future-month input is outside this screen's scope; retain its existing
        # validation requirement without reading or inventing other plan months.
        destinations=tuple(DestinationValidationInput(d.id, d.name, d.display_order, d.required_for_report, actuals[d.id].quantity_ea if d.id in actuals else None, None) for d in active),
        current_import_batches=tuple(ImportBatchInput(b.batch_id, b.report_month, b.source_type, b.file_sha256, b.source_filename, b.is_current) for b in batches),
    )
    issues.extend(validate_report(context).issues)
    if not historical["cost"].comparison_year.complete:
        issues.append(ValidationIssue("MISSING_PRIOR_YEAR_COST", "error", "전년도 운반비 이력이 누락되었습니다: " + ", ".join(historical["cost"].comparison_year.missing_months)))
    if not active:
        issues.append(ValidationIssue("NO_DESTINATIONS", "error", "활성 납품처를 등록하세요."))
    groups = [(group, calculate_group(calculations, masters.group_members(group.id), group_id=group.id)) for group in masters.list_groups() if group.active]
    required = [d for d in active if d.required_for_report]
    total = calculate_total(calculations, active)
    # The generation gate is the same snapshot preflight used by the exporter.
    # Partial preview rows remain available while users correct missing inputs.
    try:
        bundle, _ = ReportService(db, RuntimePaths.from_root(Path(__file__).resolve().parents[2])).prepare(month)
        issues = list(bundle.validation.issues)
        by_id = {d.id: d for d in bundle.destinations}
        group_by_id = {g.id: g for g in bundle.groups}
        rows, groups = [], []
        for row in bundle.report.rows:
            if row is None:
                continue
            result = row.calculation
            if getattr(result, "destination_id", None) is not None:
                rows.append({"destination": by_id[result.destination_id], "result": result, "actual_quantity": result.actual_quantity, "actual_cost": result.actual_cost_won})
            else:
                groups.append((group_by_id[result.group_id], result))
        total = bundle.report.total.calculation
        historical = {
            "quantity": historical_averages(month, bundle.report.total.quantity_by_month),
            "cost": historical_averages(month, bundle.report.total.cost_won_by_month, value_kind=HistoricalValueKind.MONEY),
        }
    except ReportGenerationError as error:
        issues = list(error.issues)
    return {"rows": rows, "groups": groups, "total": total, "issues": issues, "can_generate": ValidationResult(tuple(issues)).can_generate,
            "batches": batches, "plan_count": sum(d.id in plans for d in required), "actual_count": sum(d.id in actuals for d in required), "required_count": len(required), "sales": sales, "historical": historical}


@router.get("/")
def dashboard(request: Request, month: str = MONTH):
    return render(request, "dashboard.html", **report_preview_data(database(request), month))


@router.get("/months/{month}/preview")
def preview(request: Request, month: str):
    return render(request, "report_preview.html", **report_preview_data(database(request), month))


@router.post("/months/{month}/generate/{kind}")
async def generate(request: Request, month: str, kind: str):
    await form(request, "generate")
    require_month(month)
    if kind not in {"excel", "ppt"}:
        raise FormError("지원하지 않는 생성 종류입니다.")
    try:
        result = ReportService(database(request), request.app.state.runtime_paths).generate(month)
    except ReportGenerationError as error:
        raise FormError(str(error), 409) from error
    filename = "review.xlsx" if kind == "excel" else "report.pptx"
    return FileResponse(result.directory / filename, filename=f"{month}-{filename}")


@router.get("/destinations")
def destinations(request: Request):
    repo = MasterRepository(database(request))
    return render(request, "destinations.html", destinations=repo.list_destinations(), aliases=repo.list_aliases())


@router.post("/destinations/{destination_id}")
async def save_destination(request: Request, destination_id: int):
    data = await form(request, "save_destination")
    repo = MasterRepository(database(request))
    values = {key: data.get(key) == "on" for key in ("active", "required_for_report", "include_quantity_total", "include_cost_total", "include_sales_total")}
    values.update(name=data.get("name", ""), display_order=integer(data.get("display_order", ""), "표시 순서"), representative_item=data.get("representative_item", "").strip() or None)
    if destination_id == 0:
        repo.create_destination(**values)
    else:
        repo.update_destination(destination_id, **values)
    return preview_redirect()


@router.post("/destinations/{destination_id}/aliases")
async def save_alias(request: Request, destination_id: int):
    data = await form(request, "save_alias")
    MasterRepository(database(request)).add_alias(destination_id, data.get("raw_name", ""), data.get("source_type", ""))
    return preview_redirect()


@router.post("/aliases/{alias_id}/remove")
async def remove_alias(request: Request, alias_id: int):
    await form(request, "remove_alias")
    MasterRepository(database(request)).remove_alias(alias_id)
    return preview_redirect()


@router.get("/groups")
def groups(request: Request):
    repo = MasterRepository(database(request))
    return render(request, "groups.html", groups=repo.list_groups(), destinations=repo.list_destinations(), members={group.id: repo.group_members(group.id) for group in repo.list_groups()})


@router.post("/groups/{group_id}")
async def save_group(request: Request, group_id: int):
    data = await form(request, "save_group")
    ids = [integer(value, "그룹 납품처") for value in data.getlist("member_id") if value.strip()]
    ordered = [(integer(data.get(f"order_{item}", str(index)), "구성 순서"), index, item) for index, item in enumerate(ids)]
    ids = [item for _, _, item in sorted(ordered)]
    quantity_ids = {integer(value, "수량 포함") for value in data.getlist("quantity_ids")}
    cost_ids = {integer(value, "비용 포함") for value in data.getlist("cost_ids")}
    if not quantity_ids.issubset(ids) or not cost_ids.issubset(ids):
        raise FormError("그룹 구성원만 수량·비용에 포함할 수 있습니다.")
    MasterRepository(database(request)).save_group_form(group_id or None, data.get("name", ""), integer(data.get("display_order", ""), "표시 순서"), data.get("active") == "on", [(item, item in quantity_ids, item in cost_ids) for item in ids])
    return preview_redirect()


@router.get("/months/{month}/inputs")
def monthly_inputs(request: Request, month: str):
    require_month(month)
    db = database(request)
    snapshot = MonthlyInputRepository(db).get_form_snapshot(month)
    return render(
        request, "monthly_inputs.html", destinations=snapshot.destinations,
        plans={p.destination_id: p for p in snapshot.plans},
        actuals={a.destination_id: a for a in snapshot.actuals}, sales=snapshot.sales,
        monthly_revision=snapshot.monthly_revision,
    )


@router.post("/months/{month}/inputs")
async def save_inputs(request: Request, month: str):
    require_month(month)
    data = await form(request, "save_inputs")
    if len(data.getlist("monthly_revision")) != 1 or not data.get("monthly_revision"):
        raise FormError("월 입력 검토 정보가 없습니다. 화면을 다시 열어 검토한 뒤 저장하세요.", 409)
    keys = ("destination_id", "plan_quantity", "plan_cost", "representative_item", "actual_quantity", "source_note")
    columns = [data.getlist(key) for key in keys]
    if len({len(column) for column in columns}) != 1:
        raise FormError("납품처별 입력 행이 누락되었습니다. 화면을 다시 여세요.")
    rows = []
    for values in zip(*columns):
        row = dict(zip(keys, values))
        row.update(destination_id=integer(row["destination_id"], "납품처"), plan_quantity=quantity(row["plan_quantity"], "계획 수량"), plan_cost=integer(row["plan_cost"], "계획 비용", blank=True), actual_quantity=quantity(row["actual_quantity"], "실적 수량"))
        rows.append(row)
    amount = integer(data.get("sales", ""), "매출", blank=True)
    confirmed = data.get("confirmed_at", "").strip()
    if (amount is None) != (not confirmed):
        raise FormError("매출액과 확정일을 함께 입력하세요.")
    sales = None if amount is None else MonthlySales(month, amount, data.get("sales_source_note", "").strip() or None, confirmed)
    MonthlyInputRepository(database(request)).save_form(
        month, rows, sales, expected_revision=data["monthly_revision"]
    )
    return preview_redirect()


def current_review(request: Request):
    review = request.app.state.import_reviews.get(request.cookies.get("csrf_token"))
    if review is None:
        raise FormError("먼저 원본 파일의 드라이런을 실행하세요.")
    return review


@router.get("/imports/review")
def import_review(request: Request):
    review = request.app.state.import_reviews.get(request.cookies.get("csrf_token"))
    if review is not None:
        review = LegacyMigrationService(database(request)).dry_run(review.source_path)
        request.app.state.import_reviews[request.cookies["csrf_token"]] = review
    return render(request, "import_review.html", review=review, destinations=MasterRepository(database(request)).list_destinations())


@router.post("/imports/dry-run")
async def dry_run(request: Request):
    data = await form(request, "dry_run")
    review = LegacyMigrationService(database(request)).dry_run(data.get("source_path", ""))
    request.app.state.import_reviews[request.cookies["csrf_token"]] = review
    return RedirectResponse("/imports/review", status_code=303)


def require_revision(data: FormData, review) -> None:
    for name in ("source_sha256", "master_revision", "database_revision", "revision"):
        if data.get(name) != getattr(review, name):
            raise FormError("검토 이후 자료가 변경되었습니다. 드라이런을 다시 실행하세요.", 409)


@router.post("/imports/aliases")
async def resolve_import_alias(request: Request):
    data = await form(request, "resolve_alias")
    old_review = current_review(request)
    review = LegacyMigrationService(database(request)).dry_run(old_review.source_path)
    require_revision(data, review)
    alias, source_type = data.get("alias", ""), data.get("source_type", "")
    if not any(item.alias == alias and item.source_type == source_type for item in review.unknown_aliases):
        raise FormError("현재 드라이런의 미등록 명칭만 연결할 수 있습니다.")
    repo = MasterRepository(database(request))
    if data.get("action") == "link":
        repo.add_alias(integer(data.get("destination_id", ""), "납품처"), alias, source_type)
    elif data.get("action") == "create":
        repo.create_destination(data.get("name", ""), integer(data.get("display_order", ""), "표시 순서"),
            active=data.get("active") == "on", required_for_report=data.get("required_for_report") == "on",
            representative_item=data.get("representative_item", "").strip() or None,
            include_quantity_total=data.get("include_quantity_total") == "on", include_cost_total=data.get("include_cost_total") == "on", include_sales_total=data.get("include_sales_total") == "on", initial_alias=(alias, source_type))
    else:
        raise FormError("기존 연결 또는 새 등록을 선택하세요.")
    request.app.state.import_reviews[request.cookies["csrf_token"]] = LegacyMigrationService(database(request)).dry_run(review.source_path)
    return RedirectResponse("/imports/review", status_code=303)


@router.post("/imports/confirm")
async def confirm_import(request: Request):
    data = await form(request, "confirm_import")
    review = current_review(request)
    if data.get("confirmed") != "yes":
        raise FormError("검토 결과를 확인하고 가져오기 확정에 체크하세요.")
    require_revision(data, review)
    LegacyMigrationService(database(request)).commit(
        review.source_path, report_month=MONTH, confirmed=True,
        **{f"expected_{name}": data[name] for name in ("source_sha256", "master_revision", "database_revision", "revision")},
    )
    request.app.state.import_reviews.pop(request.cookies["csrf_token"], None)
    return preview_redirect()


def register_web(app: FastAPI) -> None:
    async def handle_error(request: Request, error: Exception):
        status = error.status if isinstance(error, FormError) else 409 if isinstance(error, (LegacyMigrationRevisionError, MonthlyInputRevisionError)) else 422
        message = str(error) if isinstance(error, (FormError, LegacyMigrationError, MonthlyInputRevisionError)) else f"입력 내용을 확인하세요. 중복 명칭·납품처·숫자·확정일 또는 월 잠금으로 저장하지 못했습니다. ({error})"
        return render(request, "base.html", status=status, error=message)

    for error_type in (FormError, MasterDataError, MonthlyInputError, LegacyMigrationError, sqlite3.IntegrityError):
        app.add_exception_handler(error_type, handle_error)
    app.include_router(router)
    from app.web.report_input_routes import router as report_input_router
    app.include_router(report_input_router)
