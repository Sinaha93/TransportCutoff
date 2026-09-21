"""Snapshot-based August reporting; rendering never holds a live write lock.

The input revision hashes an immutable, canonical dump of every application
input table (including master rules, selected batches, history and locks).
Concurrent edits are allowed after snapshot capture and affect the next run.
The audit retains that dump and the exact typed report model, not just a pointer
to mutable rows. Files are published together, then audited in a short SQLite
transaction; any raised failure removes only this run's unpublished files.
"""
from __future__ import annotations

from dataclasses import dataclass, fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile

from PIL import Image
from pptx import Presentation

from app.config import RuntimePaths
from app.db import Database
from app.domain.calculations import calculate_destination, calculate_group, calculate_total, historical_averages, HistoricalValueKind
from app.domain.models import Destination, DestinationAlias, ReportGroup, GroupMember
from app.domain.validation import ValidationContext, ValidationIssue, validate_report, SalesInput, ReportMonthSources, DestinationValidationInput, NextMonthPlanInput, ImportBatchInput, UnresolvedDestinationAlias
from app.reporting import charts, pptx_report, xlsx_review
from app.reporting.pptx_report import ReportRow, SalesReport, PlanReport, PptReport
from app.reporting.xlsx_review import ReviewWorkbookData, ImportBatchEvidence, OperationEvidence
from tools.verify_ppt_layout import verify_ppt_layout


class ReportGenerationError(ValueError):
    def __init__(self, issues):
        self.issues = tuple(issues)
        super().__init__("\n".join(issue.message for issue in self.issues))


@dataclass(frozen=True)
class ReportRun:
    run_id: int
    directory: Path
    input_revision: str


def _json_value(value):
    if isinstance(value, Decimal):
        return {"$decimal": str(value)}
    if isinstance(value, (date, datetime)):
        return {"$datetime" if isinstance(value, datetime) else "$date": value.isoformat()}
    if isinstance(value, Enum):
        return value.value
    if is_dataclass(value):
        return {field.name: _json_value(getattr(value, field.name)) for field in fields(value)}
    if isinstance(value, dict):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [_json_value(item) for item in value]
    return value


def canonical_json(value) -> str:
    return json.dumps(_json_value(value), ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def _issue(code, message, destination_id=None):
    return ValidationIssue(code, "error", message, destination_id)


def _decimal(value):
    return None if value is None else Decimal(value)


def _filename(value):
    return str(value).replace("\\", "/").rsplit("/", 1)[-1]


class ReportService:
    def __init__(self, database: Database, paths: RuntimePaths, *, clock=datetime.now):
        self.database, self.paths, self.clock = database, paths, clock

    def _snapshot(self):
        # Online backup establishes one SQLite snapshot without relying on
        # independently opened repository connections sharing a transaction.
        snapshot = sqlite3.connect(":memory:")
        snapshot.row_factory = sqlite3.Row
        try:
            with self.database.connection() as live:
                live.backup(snapshot)
            names = [r[0] for r in snapshot.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name <> 'report_runs' ORDER BY name")]
            return {name: sorted((dict(row) for row in snapshot.execute('SELECT * FROM "' + name.replace('"', '""') + '"')), key=canonical_json) for name in names}
        finally:
            snapshot.close()

    def prepare(self, month="2026-08"):
        """Public read-only preflight. No output folders or audit rows are made."""
        if month != "2026-08":
            raise ReportGenerationError((_issue("REPORT_MONTH_MISMATCH", "보고서 생성 범위는 2026-08입니다. 9월 계획은 다음 달 계획에 입력하세요."),))
        try:
            snapshot = self._snapshot()
            bundle = self._build(snapshot, self.clock().date())
            xlsx_review._validate_report_identity(bundle)
            for builder in (charts.build_quantity_chart_data, charts.build_cost_chart_data, charts.build_combined_chart_data):
                builder(bundle.report.charts)
        except ReportGenerationError:
            raise
        except (sqlite3.Error, OSError) as error:
            raise ReportGenerationError((_issue("REPORT_INPUT_UNAVAILABLE", "보고 입력을 읽지 못했습니다. 데이터베이스 사용 상태를 확인하고 다시 시도하세요."),)) from error
        except (ValueError, TypeError, KeyError) as error:
            raise ReportGenerationError((_issue("REPORT_MODEL_INVALID", "보고서 입력·마스터·이력 간 값이 일치하지 않습니다. 검토용 입력과 12개 납품처/집계그룹을 확인하세요."),)) from error
        return bundle, canonical_json(snapshot)

    def generate(self, month="2026-08", *, requested_run_id=None):
        stamp = self.clock()
        bundle, snapshot_json = self.prepare(month)
        model_json = canonical_json(bundle)
        revision = sha256(snapshot_json.encode()).hexdigest()
        name = requested_run_id or "run-" + stamp.strftime("%Y%m%d-%H%M%S")
        if not re.fullmatch(r"run-[0-9]{8}-[0-9]{6}(?:-[0-9]{3,})?", name):
            raise ReportGenerationError((_issue("INVALID_RUN_ID", "실행 이름 형식이 올바르지 않습니다. 보고서를 다시 생성하세요."),))
        parent = self.paths.outputs / month
        parent.mkdir(parents=True, exist_ok=True)
        temporary = Path(tempfile.mkdtemp(prefix=".run-", dir=parent))
        final = None
        try:
            self._charts(bundle, temporary)
            self._xlsx(bundle, temporary)
            self._pptx(bundle, temporary)
            self._verify(bundle, temporary)
            final = self._rename(temporary, parent, name, requested_run_id is not None)
            files = [{"path": p.relative_to(self.paths.outputs).as_posix(), "sha256": sha256(p.read_bytes()).hexdigest(), "size": p.stat().st_size} for p in sorted(final.iterdir())]
            manifest = {"format_version": 1, "files": files, "input_snapshot_json": snapshot_json, "model_json": model_json, "model_sha256": sha256(model_json.encode()).hexdigest()}
            run_id = self._audit(month, revision, manifest, stamp)
            return ReportRun(run_id, final, revision)
        except Exception as error:
            if final is not None:
                shutil.rmtree(final)
            raise ReportGenerationError((_issue("REPORT_GENERATION_FAILED", "보고서 생성을 완료하지 못했습니다. 저장 공간·파일 사용 여부를 확인하고 다시 생성하세요. 기존 보고서는 보존되었습니다."),)) from error
        finally:
            if temporary.exists():
                shutil.rmtree(temporary)

    def _charts(self, bundle, directory):
        for name, renderer in (("quantity", charts.render_quantity_chart), ("cost", charts.render_cost_chart), ("combined", charts.render_combined_chart)):
            renderer(bundle.report.charts, directory / (name + ".png"))

    def _xlsx(self, bundle, directory):
        xlsx_review.export_review_workbook(bundle, directory / "review.xlsx")

    def _pptx(self, bundle, directory):
        pptx_report.generate_pptx(self.paths.template, bundle.report, directory / "report.pptx")

    def _verify(self, bundle, directory):
        if {p.name for p in directory.iterdir()} != {"quantity.png", "cost.png", "combined.png", "review.xlsx", "report.pptx"}:
            raise ValueError("unexpected output set")
        for name, size in (("quantity", charts.QUANTITY_CANVAS_PIXELS), ("cost", charts.COST_CANVAS_PIXELS), ("combined", charts.COMBINED_CANVAS_PIXELS)):
            with Image.open(directory / (name + ".png")) as image:
                if image.format != "PNG" or image.size != size:
                    raise ValueError("invalid chart")
                image.verify()
        xlsx_review._validate_saved_workbook(directory / "review.xlsx", bundle)
        if verify_ppt_layout(self.paths.template, directory / "report.pptx"):
            raise ValueError("PowerPoint layout verification failed")
        if len(Presentation(directory / "report.pptx").slides) != 6:
            raise ValueError("invalid PowerPoint")

    @staticmethod
    def _rename(temporary, parent, name, exact):
        index = 0
        while True:
            final = parent / (name if index == 0 else f"{name}-{index:03}")
            if not final.exists():
                try:
                    os.rename(temporary, final)
                    return final
                except FileExistsError:
                    pass
            if exact:
                raise FileExistsError("requested run already exists")
            index += 1

    def _audit(self, month, revision, manifest, stamp):
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                row = connection.execute("INSERT INTO report_runs(report_month,status,input_revision,output_hashes_json,created_at) VALUES (?,'completed',?,?,?)", (month, revision, canonical_json(manifest), stamp.isoformat())).lastrowid
                connection.commit()
                return row
            except Exception:
                connection.rollback()
                raise

    def _build(self, snapshot, created_on):
        month, next_month = "2026-08", "2026-09"
        destinations = tuple(sorted((Destination(**{**row, **{key: bool(row[key]) for key in ("active", "required_for_report", "include_quantity_total", "include_cost_total", "include_sales_total")}}) for row in snapshot["destinations"]), key=lambda d: (d.display_order, d.id)))
        active = tuple(d for d in destinations if d.active and d.required_for_report)
        aliases = tuple(DestinationAlias(**r) for r in snapshot["destination_aliases"])
        groups = tuple(sorted((ReportGroup(**{**r, "active": bool(r["active"])}) for r in snapshot["report_groups"]), key=lambda g: (g.display_order, g.id)))
        names = {d.id: d.name for d in destinations}
        members = tuple(GroupMember(**{**r, "name": names[r["destination_id"]], "include_quantity": bool(r["include_quantity"]), "include_cost": bool(r["include_cost"])}) for r in snapshot["report_group_members"])
        plans = {(r["report_month"], r["destination_id"]): r for r in snapshot["monthly_plans"]}
        quantities = {(r["report_month"], r["destination_id"]): _decimal(r["quantity_ea_text"]) for r in snapshot["monthly_actual_quantities"]}
        costs = {(r["report_month"], r["destination_id"]): r["cost_won"] for r in snapshot["monthly_actual_costs"]}
        sales = {r["report_month"]: r for r in snapshot["monthly_sales"]}
        supplementary = {r["report_month"]: r for r in snapshot["report_supplemental_inputs"]}
        current_batches = {r["id"]: r for r in snapshot["import_batches"] if r["is_current"]}
        entries = [r for r in snapshot["transport_entries"] if r["import_batch_id"] in current_batches]
        entry_groups = {}
        for row in entries:
            if row["destination_id"] is not None:
                entry_groups.setdefault((row["report_month"], row["destination_id"]), []).append(row)
        costs.update({key: sum(r["cost_won"] for r in rows) for key, rows in entry_groups.items()})
        aliases_by_source = {(a.source_type, a.raw_name): a.destination_id for a in aliases}
        unresolved = []
        for row in entries:
            if row["report_month"] != month:
                continue
            batch = current_batches[row["import_batch_id"]]
            alias = row["source_alias"] or row["unresolved_alias"]
            if row["destination_id"] is None or alias is None or aliases_by_source.get((batch["source_type"], alias)) != row["destination_id"]:
                unresolved.append(UnresolvedDestinationAlias(alias or "원천 별칭 미기록", f'{_filename(batch["source_filename"])}: {row["source_sheet"]}!{row["source_row"]}'))
        months = sorted({m for m, _ in plans} | {m for m, _ in quantities} | {m for m, _ in costs})
        calculations = {}
        totals = {}
        for m in months:
            by_destination = {}
            for d in active:
                plan = plans.get((m, d.id))
                q, c = quantities.get((m, d.id)), costs.get((m, d.id))
                by_destination[d.id] = calculate_destination(_decimal(plan["quantity_ea_text"]) if plan else None, plan["cost_won"] if plan else None, q if c is not None else None, c if q is not None else None, destination_id=d.id)
            calculations[m] = by_destination
            totals[m] = calculate_total(by_destination, active)
        total_q = {m: t.actual_quantity for m, t in totals.items() if m <= month}
        total_c = {m: t.actual_cost_won for m, t in totals.items() if m <= month}
        batch_rows = sorted((r for r in current_batches.values() if r["report_month"] == month), key=lambda r: r["id"])
        sales_row = sales.get(month)
        context = ValidationContext(month, SalesInput(sales_row["amount_won"], sales_row["confirmed_at"], "월 매출 확정 입력") if sales_row else None, historical_averages(month, total_q).comparison_year, ReportMonthSources(month, month, month, month, month), tuple(unresolved), tuple(DestinationValidationInput(d.id, d.name, d.display_order, d.required_for_report, quantities.get((month, d.id)), NextMonthPlanInput(next_month, _decimal(plans[next_month, d.id]["quantity_ea_text"])) if (next_month, d.id) in plans else None) for d in active), tuple(ImportBatchInput(r["id"], month, r["source_type"], r["file_sha256"], _filename(r["source_filename"]), True) for r in batch_rows))
        issues = list(validate_report(context).issues)
        if len(active) != 12 or sum(g.active for g in groups) > 2:
            issues.append(_issue("REPORT_LAYOUT_CAPACITY", "현재 보고서는 필수 활성 납품처 12개와 집계그룹 최대 2개를 사용합니다. 마스터 표시 구성을 확인하세요."))
        if not batch_rows:
            issues.append(_issue("MISSING_IMPORT_BATCH", "8월 운반비 원천 자료를 가져오고 현재 배치를 선택하세요."))
        for d in active:
            for missing, code, label in (((month, d.id) not in plans, "MISSING_CURRENT_PLAN", "8월 계획"), ((month, d.id) not in costs, "MISSING_ACTUAL_COST", "8월 운반비")):
                if missing:
                    issues.append(_issue(code, f"{d.name}: {label}을 입력하세요.", d.id))
        for m in (month, next_month):
            supplement = supplementary.get(m, {})
            required = ["planned_sales_won", "nonregular_planned_quantity_text", "nonregular_planned_cost_won"]
            if m == month:
                required += ["nonregular_actual_quantity_text", "nonregular_actual_cost_won"]
            if any(supplement.get(field) is None for field in required):
                issues.append(_issue("MISSING_REPORT_SUPPLEMENT", f"{m}: 매출 계획과 비정규 수량·운반비를 입력하세요. 해당 없음은 확인한 0을 입력하세요."))
        for values, label, kind in ((total_q, "실적 수량", HistoricalValueKind.QUANTITY), (total_c, "실적 운반비", HistoricalValueKind.MONEY), ({m: t.planned_quantity for m, t in totals.items()}, "계획 수량", HistoricalValueKind.QUANTITY), ({m: t.planned_cost_won for m, t in totals.items()}, "계획 운반비", HistoricalValueKind.MONEY)):
            history = historical_averages(month, values, value_kind=kind)
            periods = (history.twelve_month, history.comparison_year) if label.startswith("실적") else (history.twelve_month,)
            missing = sorted({m for period in periods for m in period.missing_months})
            if missing:
                issues.append(_issue("MISSING_REPORT_HISTORY", f"{label} 이력이 누락되었습니다. 해당 월 자료를 입력하세요: " + ", ".join(missing)))
        if any(issue.severity == "error" for issue in issues):
            raise ReportGenerationError(issues)
        def direct_row(d, m):
            return ReportRow(str(d.id), d.name, calculations[m][d.id], {hm: quantities.get((hm, d.id)) for hm in months if hm <= month}, {hm: costs.get((hm, d.id)) for hm in months if hm <= month})
        def total_row(m):
            return ReportRow("total", "합계", totals[m], total_q, total_c)
        def nonregular_row(m):
            r = supplementary[m]
            current = m == month
            calc = calculate_destination(_decimal(r["nonregular_planned_quantity_text"]), r["nonregular_planned_cost_won"], _decimal(r["nonregular_actual_quantity_text"]) if current else None, r["nonregular_actual_cost_won"] if current else None)
            return ReportRow("nonregular", "비정규 운반비", calc, {hm: _decimal(v["nonregular_actual_quantity_text"]) for hm, v in supplementary.items() if hm <= month}, {hm: v["nonregular_actual_cost_won"] for hm, v in supplementary.items() if hm <= month})
        current_rows = [direct_row(d, month) for d in active]
        for g in groups:
            if not g.active:
                continue
            rules = sorted((r for r in members if r.group_id == g.id), key=lambda r: (r.display_order, r.destination_id))
            values = {m: calculate_group(calculations[m], rules, group_id=g.id) for m in months if m <= month}
            current_rows.append(ReportRow(f"group:{g.id}", g.name, values[month], {m: v.actual_quantity for m, v in values.items()}, {m: v.actual_cost_won for m, v in values.items()}))
        current_rows.extend([None] * (14 - len(current_rows)))
        sales_history = {m: r["amount_won"] for m, r in sales.items() if m <= month}
        chart = charts.ChartReport(month, {m: t.planned_quantity for m, t in totals.items() if m <= month}, total_q, {m: t.planned_cost_won for m, t in totals.items() if m <= month}, total_c)
        report = PptReport(month, created_on, tuple(current_rows), nonregular_row(month), total_row(month), SalesReport(supplementary[month]["planned_sales_won"], sales_row["amount_won"], sales_history), PlanReport(next_month, tuple(direct_row(d, next_month) for d in active), nonregular_row(next_month), total_row(next_month), SalesReport(supplementary[next_month]["planned_sales_won"], None, sales_history)), chart)
        evidence_batches = tuple(ImportBatchEvidence(r["id"], f'batch:{r["id"]}', month, r["source_type"], _filename(r["source_filename"]), r["file_sha256"], datetime.fromisoformat(r["imported_at"])) for r in batch_rows)
        operations = []
        def operation(kind, role, m, destination_id, label, quantity, cost, locator):
            operations.append(OperationEvidence(kind, role, "manual", None, f"{kind}:{role}:{m}:{destination_id}", m, None, destination_id, label, quantity, cost, locator))
        for period in (report, report.next_month):
            m = month if period is report else next_month
            for row in period.rows:
                if row is None or getattr(row.calculation, "destination_id", None) is None:
                    continue
                d = row.calculation.destination_id
                c = row.calculation
                operation("destination", "plan", m, d, row.label, c.planned_quantity, c.planned_cost_won, "월 계획 입력")
                if period is report:
                    sources = entry_groups.get((m, d), [])
                    if sources:
                        locator = "ERP 수량 입력 + 현재 운반비 배치 합계: " + ", ".join(f'{r["import_batch_id"]}:{r["source_sheet"]}!{r["source_row"]}' for r in sources)
                        operations.append(OperationEvidence("destination", "actual", "derived", None, f"aggregate:{m}:{d}", m, None, d, row.label, c.actual_quantity, c.actual_cost_won, locator))
                    else:
                        operation("destination", "actual", m, d, row.label, c.actual_quantity, c.actual_cost_won, "월 실적 수량·운반비 확정 입력")
            extra = period.nonregular.calculation
            operation("nonregular", "plan", m, None, report.nonregular.label, extra.planned_quantity, extra.planned_cost_won, "비정규 계획 입력")
            operation("sales", "plan", m, None, "매출액", None, period.sales.planned_won, "매출 계획 입력")
            if period is report:
                operation("nonregular", "actual", m, None, report.nonregular.label, extra.actual_quantity, extra.actual_cost_won, "비정규 실적 입력")
                operation("sales", "actual", m, None, "매출액", None, report.sales.actual_won, "매출 확정 입력")
        return ReviewWorkbookData(report, context, destinations, aliases, groups, members, evidence_batches, tuple(operations))
