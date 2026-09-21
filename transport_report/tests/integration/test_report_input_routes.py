from decimal import Decimal

from tests.integration.test_web_routes import database, masters, client, post, hidden_fields
from app.repositories.monthly_inputs import MonthlyInputRepository
from app.repositories.report_inputs import ReportInputRepository, ReportSupplementalInputs


PATH = "/months/2026-08/report-inputs"


def payload(client, masters, **changes):
    fields = hidden_fields(client.get(PATH).text)
    values = dict(intent="save_report_inputs", supplemental_revision=fields.get("supplemental_revision", "missing"),
        current_planned_sales="", current_nonregular_planned_quantity="", current_nonregular_planned_cost="",
        current_nonregular_actual_quantity="0", current_nonregular_actual_cost="0", current_source_note="8월 확인",
        next_planned_sales="0", next_nonregular_planned_quantity="0", next_nonregular_planned_cost="0", next_source_note="9월 계획")
    values.update(changes)
    pairs = list(values.items())
    for d in masters.list_destinations():
        pairs.extend([("destination_id", str(d.id)), ("plan_quantity", "0"), ("plan_cost", "0"), ("representative_item", "품목")])
    return pairs


def test_current_and_explicit_next_plan_save_and_clear(client, masters, database):
    assert PATH in client.get("/months/2026-08/inputs").text
    page = client.get(PATH)
    assert page.status_code == 200
    assert "2026-09" in page.text
    assert post(client, PATH, payload(client, masters)).status_code == 303
    repo = ReportInputRepository(database)
    assert repo.get("2026-08").nonregular_actual_quantity_ea == Decimal(0)
    assert repo.get("2026-09").planned_sales_won == 0
    monthly = MonthlyInputRepository(database)
    assert monthly.get_plan("2026-09", 1).cost_won == 0
    assert monthly.get_sales("2026-08") is None
    assert monthly.get_sales("2026-09") is None
    assert post(client, PATH, payload(client, masters, current_nonregular_actual_quantity="", current_nonregular_actual_cost="", next_planned_sales="")).status_code == 303
    assert repo.get("2026-08").nonregular_actual_cost_won is None
    assert repo.get("2026-09").planned_sales_won is None
    assert client.get("/months/2026-09/inputs").status_code == 422
    assert client.get("/months/2026-09/report-inputs").status_code == 422


def test_pair_error_csrf_and_locked_next_month_are_atomic(client, masters, database):
    data = payload(client, masters, current_nonregular_actual_cost="")
    assert post(client, PATH, data).status_code == 422
    assert ReportInputRepository(database).get("2026-08") is None
    data = payload(client, masters)
    assert post(client, PATH, data, csrf=False).status_code == 403
    MonthlyInputRepository(database).finalize_month("2026-09", "locked")
    assert post(client, PATH, data).status_code == 422
    assert ReportInputRepository(database).get("2026-08") is None


def test_stale_changes_to_next_plans_or_supplemental_rejected(client, masters, database):
    data = payload(client, masters)
    MonthlyInputRepository(database).save_plan("2026-09", 1, 1, 100, "new")
    response = post(client, PATH, data)
    assert response.status_code == 409
    assert "다시 열어" in response.text
    assert ReportInputRepository(database).get("2026-08") is None
    data = payload(client, masters)
    ReportInputRepository(database).save(ReportSupplementalInputs("2026-08", planned_sales_won=10))
    assert post(client, PATH, data).status_code == 409
    assert ReportInputRepository(database).get("2026-08").planned_sales_won == 10


def test_next_plan_screen_preserves_existing_next_actuals(client, masters, database):
    repo = ReportInputRepository(database)
    repo.save(ReportSupplementalInputs("2026-09", nonregular_actual_quantity_ea=Decimal(7), nonregular_actual_cost_won=700))
    assert post(client, PATH, payload(client, masters)).status_code == 303
    assert repo.get("2026-09").nonregular_actual_quantity_ea == 7
    assert repo.get("2026-09").nonregular_actual_cost_won == 700


def test_missing_and_duplicate_supplemental_revision_rejected(client, masters):
    data = payload(client, masters)
    assert post(client, PATH, [(k, v) for k, v in data if k != "supplemental_revision"]).status_code == 409
    assert post(client, PATH, data + [("supplemental_revision", "duplicate")]).status_code == 409
    assert post(client, PATH, data + [("next_planned_sales", "999")]).status_code == 422
