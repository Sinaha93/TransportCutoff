from __future__ import annotations

import asyncio
import hashlib
import re
from decimal import Decimal
from urllib.parse import urlencode

import pytest
import httpx

from app import main
from app.db import Database
from app.repositories.masters import MasterRepository
from app.repositories.monthly_inputs import MonthlyInputRepository
from tests.fixtures.build_legacy_fixture import DESTINATIONS, build_legacy_fixture


@pytest.fixture
def database(tmp_path):
    database = Database(tmp_path / "web.sqlite3")
    database.migrate()
    return database


@pytest.fixture
def masters(database):
    repo = MasterRepository(database)
    for order, name in enumerate(DESTINATIONS, 1):
        destination = repo.create_destination(name, order, include_quantity_total=name != "당진")
        repo.add_alias(destination.id, name, "legacy_workbook")
    return repo


@pytest.fixture
def client(database, masters):
    client = WebClient(main.create_app(database=database))
    client.get("/")
    return client


class WebClient:
    """Use the project's httpx ASGI stack without Starlette's optional httpx2."""

    def __init__(self, app):
        self.app = app
        self.cookies = httpx.Cookies()

    def request(self, method, path, **kwargs):
        async def send():
            async with httpx.AsyncClient(transport=httpx.ASGITransport(app=self.app), base_url="http://testserver", cookies=self.cookies) as client:
                response = await client.request(method, path, **kwargs)
                self.cookies.update(client.cookies)
                return response
        return asyncio.run(send())

    def get(self, path, **kwargs):
        return self.request("GET", path, **kwargs)

    def post(self, path, **kwargs):
        return self.request("POST", path, **kwargs)


def post(client, path, data, *, csrf=True, revision=True):
    pairs = list(data.items()) if isinstance(data, dict) else list(data)
    if revision and path.endswith("/inputs") and not any(key == "monthly_revision" for key, _ in pairs):
        fields = hidden_fields(client.get(path).text)
        if "monthly_revision" in fields:
            pairs.append(("monthly_revision", fields["monthly_revision"]))
    if csrf:
        pairs.append(("csrf_token", client.cookies["csrf_token"]))
    return client.post(path, content=urlencode(pairs), headers={"Content-Type": "application/x-www-form-urlencoded"}, follow_redirects=False)


def monthly_form(masters, actual="0"):
    pairs = [("intent", "save_inputs"), ("sales", "0"), ("confirmed_at", "2026-08-31"), ("sales_source_note", "ERP")]
    for destination in masters.list_destinations():
        pairs.extend([("destination_id", str(destination.id)), ("plan_quantity", "0"), ("plan_cost", "0"), ("representative_item", "제품"), ("actual_quantity", actual), ("source_note", "ERP 확인")])
    return pairs


def test_monthly_input_page_distinguishes_blank_and_zero(client):
    response = client.get('/months/2026-08/inputs')
    assert response.status_code == 200
    assert 'name="actual_quantity"' in response.text
    assert 'data-zero-is-valid="true"' in response.text


def test_dashboard_order_missing_inputs_and_generation_gate(client):
    response = client.get("/")
    positions = [response.text.index(f'id="{section}"') for section in ("report-month", "import-status", "input-status", "validation-issues", "generation-actions")]
    assert positions == sorted(positions)
    assert "미입력" in response.text
    assert 'data-can-generate="false"' in response.text
    assert re.search(r'<button[^>]+disabled[^>]*>Excel 검토파일 생성', response.text)
    for kind in ("excel", "ppt"):
        assert post(client, f"/months/2026-08/generate/{kind}", {"intent": "generate"}).status_code == 409


def test_generation_rechecks_server_gate_even_if_preview_claims_valid(client, monkeypatch):
    from app.web import routes
    original = routes.report_preview_data
    def valid(database, month):
        result = original(database, month)
        result.update(can_generate=True, issues=[])
        return result
    monkeypatch.setattr(routes, "report_preview_data", valid)
    response = client.get("/")
    assert 'data-can-generate="true"' in response.text
    assert "생성 기능 준비 중" not in response.text
    assert not re.search(r'<button[^>]+disabled[^>]*>PPT 생성', response.text)
    assert post(client, "/months/2026-08/generate/ppt", {"intent": "generate"}).status_code == 409


def test_save_zero_then_blank_preserves_missing_semantics(client, masters, database):
    repo = MonthlyInputRepository(database)
    assert post(client, "/months/2026-08/inputs", monthly_form(masters)).status_code == 303
    assert repo.get_actual_quantity("2026-08", 1).quantity_ea == Decimal(0)
    assert repo.get_plan("2026-08", 1).quantity_ea == Decimal(0)
    assert repo.get_sales("2026-08").amount_won == 0
    assert post(client, "/months/2026-08/inputs", monthly_form(masters, actual="")).status_code == 303
    assert repo.get_actual_quantity("2026-08", 1) is None
    assert "실적 수량이 없습니다" in client.get("/months/2026-08/preview").text


@pytest.mark.parametrize("field,value", [("actual_quantity", "-1"), ("plan_quantity", "NaN"), ("plan_quantity", "1000000001"), ("plan_quantity", "0.00001"), ("plan_cost", "1.2"), ("plan_cost", ""), ("confirmed_at", "2026-02-31")])
def test_invalid_monthly_form_does_not_save(client, masters, database, field, value):
    data = monthly_form(masters)
    index = next(i for i, pair in enumerate(data) if pair[0] == field)
    data[index] = (field, value)
    response = post(client, "/months/2026-08/inputs", data)
    assert response.status_code == 422
    assert "입력" in response.text
    assert MonthlyInputRepository(database).list_plans("2026-08") == []


def test_monthly_form_rolls_back_all_rows_and_sales_on_database_failure(client, masters, database):
    repo = MonthlyInputRepository(database)
    repo.save_actual_quantity("2026-08", 1, 7)
    with database.connection() as connection:
        connection.execute("CREATE TRIGGER fail_sales BEFORE INSERT ON monthly_sales BEGIN SELECT RAISE(ABORT, 'test rollback'); END")
        connection.commit()
    response = post(client, "/months/2026-08/inputs", monthly_form(masters))
    assert response.status_code == 422
    assert repo.get_actual_quantity("2026-08", 1).quantity_ea == 7
    assert repo.list_plans("2026-08") == []
    assert repo.get_sales("2026-08") is None


def test_missing_duplicate_or_stale_destination_rows_rejected(client, masters):
    for data in (monthly_form(masters)[:-6], monthly_form(masters) + monthly_form(masters)[4:10]):
        assert post(client, "/months/2026-08/inputs", data).status_code == 422


def test_csrf_and_intent_are_required(client, masters):
    assert post(client, "/months/2026-08/inputs", monthly_form(masters), csrf=False).status_code == 403
    data = [(key, value) for key, value in monthly_form(masters) if key != "intent"]
    assert post(client, "/months/2026-08/inputs", data).status_code == 422


def test_destination_and_scoped_alias_edit_redirects_to_preview(client, masters):
    response = post(client, "/destinations/1", {"intent": "save_destination", "name": "새 이름", "display_order": "2", "active": "on", "required_for_report": "on", "representative_item": "부품", "include_cost_total": "on"})
    assert response.status_code == 303
    assert response.headers["location"] == "/months/2026-08/preview"
    assert masters.get_destination(1).include_quantity_total is False
    assert "새 이름" in client.get(response.headers["location"]).text
    response = post(client, "/destinations/1/aliases", {"intent": "save_alias", "raw_name": "ERP 이름", "source_type": "ERP"})
    assert response.status_code == 303
    assert masters.resolve_alias("ERP 이름", "ERP").id == 1
    assert masters.resolve_alias("ERP 이름", "transport") is None
    assert post(client, "/destinations/2/aliases", {"intent": "save_alias", "raw_name": "ERP 이름", "source_type": "ERP"}).status_code == 422


def test_group_members_order_flags_and_preview_recalculate(client, masters, database):
    group = masters.create_group("영남권", 1)
    inputs = MonthlyInputRepository(database)
    inputs.save_plan("2026-08", 1, 10, 100, None)
    inputs.save_plan("2026-08", 2, 20, 200, None)
    response = post(client, f"/groups/{group.id}", [("intent", "save_group"), ("name", "영남권 변경"), ("display_order", "1"), ("active", "on"), ("member_id", "2"), ("member_id", "1"), ("quantity_ids", "1"), ("cost_ids", "1"), ("cost_ids", "2")])
    assert response.status_code == 303
    members = masters.group_members(group.id)
    assert [member.destination_id for member in members] == [2, 1]
    assert [(member.include_quantity, member.include_cost) for member in members] == [(False, True), (True, True)]
    preview = client.get(response.headers["location"])
    assert "영남권 변경" in preview.text
    assert 'data-planned-quantity="10"' in preview.text
    assert 'data-planned-cost="300"' in preview.text
    assert post(client, f"/groups/{group.id}", {"intent": "save_group", "name": "실패", "display_order": "1", "member_id": "9999"}).status_code == 422
    assert masters.get_group(group.id).name == "영남권 변경"
    assert masters.group_members(group.id) == members


@pytest.fixture
def seeded_unknown_alias(client, tmp_path):
    source = build_legacy_fixture(tmp_path / "unknown.xlsx", unknown_alias=True)
    assert post(client, "/imports/dry-run", {"intent": "dry_run", "source_path": str(source)}).status_code == 303
    return source


def test_unknown_aliases_are_actionable(client, seeded_unknown_alias):
    response = client.get('/imports/review')
    assert '미등록 명칭' in response.text
    assert '기존 납품처에 연결' in response.text
    assert '새 납품처 등록' in response.text
    assert "누적 데이터!D2" in response.text
    assert "Mystery" in response.text


def hidden_fields(text):
    return dict(re.findall(r'<input type="hidden" name="([^"]+)" value="([^"]*)"', text))


def test_unknown_alias_link_requires_current_review_then_rechecks(client, seeded_unknown_alias, masters):
    fields = hidden_fields(client.get("/imports/review").text)
    fields.update(intent="resolve_alias", alias="Mystery", source_type="legacy_workbook", destination_id="1", action="link")
    assert post(client, "/imports/aliases", fields, csrf=False).status_code == 303
    assert masters.resolve_alias("Mystery", "legacy_workbook").id == 1
    assert "미등록 명칭: Mystery" not in client.get("/imports/review").text


def test_import_commit_requires_confirmation_and_revision_and_preserves_source(client, tmp_path, database):
    source = build_legacy_fixture(tmp_path / "legacy.xlsx")
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    assert post(client, "/imports/dry-run", {"intent": "dry_run", "source_path": str(source)}).status_code == 303
    fields = hidden_fields(client.get("/imports/review").text)
    fields["intent"] = "confirm_import"
    assert post(client, "/imports/confirm", fields, csrf=False).status_code == 422
    fields["confirmed"] = "yes"
    good_revision = fields["revision"]
    fields["revision"] = "stale"
    assert post(client, "/imports/confirm", fields, csrf=False).status_code == 409
    fields["revision"] = good_revision
    response = post(client, "/imports/confirm", fields, csrf=False)
    assert response.status_code == 303
    assert MonthlyInputRepository(database).get_plan("2026-08", 1).quantity_ea == 86
    assert hashlib.sha256(source.read_bytes()).hexdigest() == before


def test_other_month_not_in_scope_and_templates_escape_user_text(client, masters):
    assert client.get("/months/2026-09/inputs").status_code == 422
    masters.update_destination(1, name="<script>alert(1)</script>")
    response = client.get("/destinations")
    assert "&lt;script&gt;" in response.text
    assert "<script>alert(1)</script>" not in response.text
    assert client.get("/groups").status_code == 200
    assert client.get("/static/app.css").status_code == 200


def test_group_screen_order_fields_control_saved_member_order(client, masters):
    group = masters.create_group("순서 확인", 2)
    data = [("intent", "save_group"), ("name", "순서 확인"), ("display_order", "2"), ("member_id", "1"), ("member_id", "2"), ("order_1", "20"), ("order_2", "10")]
    assert post(client, f"/groups/{group.id}", data).status_code == 303
    assert [member.destination_id for member in masters.group_members(group.id)] == [2, 1]


def test_locked_month_rejects_entire_form(client, masters, database):
    repo = MonthlyInputRepository(database)
    repo.finalize_month("2026-08", "locked-revision")
    response = post(client, "/months/2026-08/inputs", monthly_form(masters))
    assert response.status_code == 422
    assert "월 잠금" in response.text
    assert repo.list_plans("2026-08") == []


def test_monthly_input_rows_only_include_active_required_destinations(client, masters):
    masters.update_destination(1, active=False)
    masters.update_destination(2, required_for_report=False)
    response = client.get("/months/2026-08/inputs")
    assert len(re.findall(r'name="actual_quantity"', response.text)) == 12
    assert 'name="destination_id" value="1"' not in response.text
    assert 'name="destination_id" value="2"' not in response.text


def test_blank_sales_clears_value_and_date_together(client, masters, database):
    data = monthly_form(masters)
    assert post(client, "/months/2026-08/inputs", data).status_code == 303
    data = [(key, "" if key in {"sales", "confirmed_at"} else value) for key, value in data]
    assert post(client, "/months/2026-08/inputs", data).status_code == 303
    assert MonthlyInputRepository(database).get_sales("2026-08") is None


def test_unknown_alias_create_is_atomic_and_rechecks(client, seeded_unknown_alias, masters, database):
    fields = hidden_fields(client.get("/imports/review").text)
    fields.update(intent="resolve_alias", alias="Mystery", source_type="legacy_workbook", action="create", name="신규 납품처", display_order="15", active="on", required_for_report="on", include_quantity_total="on", include_cost_total="on")
    assert post(client, "/imports/aliases", fields, csrf=False).status_code == 303
    destination = masters.resolve_alias("Mystery", "legacy_workbook")
    assert destination.name == "신규 납품처"
    assert destination.include_sales_total is False
    assert MonthlyInputRepository(database).list_plans("2026-08") == []


def test_alias_failure_rolls_back_new_destination(client, seeded_unknown_alias, masters, database):
    fields = hidden_fields(client.get("/imports/review").text)
    fields.update(intent="resolve_alias", alias="Mystery", source_type="legacy_workbook", action="create", name="롤백 대상", display_order="15")
    with database.connection() as connection:
        connection.execute("CREATE TRIGGER fail_alias BEFORE INSERT ON destination_aliases BEGIN SELECT RAISE(ABORT, 'alias failure'); END")
        connection.commit()
    assert post(client, "/imports/aliases", fields, csrf=False).status_code == 422
    assert all(destination.name != "롤백 대상" for destination in masters.list_destinations())


def test_changed_master_blocks_review_commit(client, tmp_path, masters, database):
    source = build_legacy_fixture(tmp_path / "changed.xlsx")
    post(client, "/imports/dry-run", {"intent": "dry_run", "source_path": str(source)})
    fields = hidden_fields(client.get("/imports/review").text)
    fields.update(intent="confirm_import", confirmed="yes")
    masters.update_destination(1, name="검토 이후 변경")
    response = post(client, "/imports/confirm", fields, csrf=False)
    assert response.status_code == 409
    assert "기준정보가 변경" in response.text
    assert MonthlyInputRepository(database).list_plans("2026-08") == []


def test_preview_current_batch_overrides_legacy_cost_and_uses_delivery_quantity(client, database, masters):
    from datetime import date
    from app.importers.hwaseong_workbook import ParsedTransportEntry
    from app.repositories.transport_entries import TransportEntryRepository
    from app.web.routes import report_preview_data

    masters.add_alias(1, "Alpha transport", "transport")
    MonthlyInputRepository(database).save_actual_quantity("2026-08", 1, 10)
    with database.connection() as connection:
        connection.execute("INSERT INTO monthly_actual_costs(report_month, destination_id, cost_won, source_type) VALUES ('2026-08', 1, 999, 'legacy_workbook')")
        connection.commit()
    repository = TransportEntryRepository(database)
    for value, digest in [(100, "a" * 64), (200, "b" * 64)]:
        row = ParsedTransportEntry(report_month="2026-08", destination_alias="Alpha transport", source_sheet="운반비", source_row=3, source_date=date(2026, 8, 1), day=1, transport_type="regular", trip_count=Decimal("2"), unit_rate_won=value // 2, cost_won=value)
        repository.import_entries(report_month="2026-08", source_filename="transport.xlsx", file_sha256=digest, rows=[row])
    values = report_preview_data(database, "2026-08")
    result = values["rows"][0]["result"]
    assert result.actual_cost_won == 200
    assert result.actual_unit_cost == Decimal("20.00")
    assert "20.00" in client.get("/months/2026-08/preview").text


def monthly_database_state(database):
    with database.connection() as connection:
        return tuple(
            tuple(tuple(row) for row in connection.execute(f"SELECT * FROM {table} ORDER BY id"))
            for table in ("monthly_plans", "monthly_actual_quantities", "monthly_sales")
        )


def test_two_tabs_cannot_overwrite_newer_monthly_values(client, masters, database):
    path = "/months/2026-08/inputs"
    tab_a = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    tab_b = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    assert tab_a == tab_b
    newer = {"plan_quantity": "4", "plan_cost": "100", "representative_item": "B 품목", "actual_quantity": "7", "source_note": "B 실적 출처", "sales": "1000", "confirmed_at": "2026-08-30", "sales_source_note": "B 매출 출처"}
    data_b = [(key, newer.get(key, value)) for key, value in monthly_form(masters)]
    data_b.append(("monthly_revision", tab_b))
    assert post(client, path, data_b).status_code == 303
    state_b = monthly_database_state(database)
    data_a = monthly_form(masters) + [("monthly_revision", tab_a)]
    response = post(client, path, data_a)
    assert response.status_code == 409
    assert "다시" in response.text and "변경" in response.text
    assert "location" not in response.headers
    assert monthly_database_state(database) == state_b
    repo = MonthlyInputRepository(database)
    assert repo.get_plan("2026-08", 1).representative_item == "B 품목"
    assert repo.get_actual_quantity("2026-08", 1).source_note == "B 실적 출처"
    assert repo.get_sales("2026-08").amount_won == 1000
    assert repo.get_sales("2026-08").confirmed_at.startswith("2026-08-30")


def test_stale_last_row_conflict_checks_revision_before_any_write(client, masters, database):
    path = "/months/2026-08/inputs"
    assert post(client, path, monthly_form(masters)).status_code == 303
    stale = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    MonthlyInputRepository(database).save_actual_quantity("2026-08", 14, 23, "새로운 마지막 행")
    before = monthly_database_state(database)
    with database.connection() as connection:
        connection.execute("CREATE TRIGGER forbid_plan_write BEFORE INSERT ON monthly_plans BEGIN SELECT RAISE(ABORT, 'revision check must precede writes'); END")
        connection.commit()
    response = post(client, path, monthly_form(masters) + [("monthly_revision", stale)])
    assert response.status_code == 409
    assert monthly_database_state(database) == before


@pytest.mark.parametrize("table,column,value", [
    ("monthly_plans", "quantity_ea_text", "8"),
    ("monthly_plans", "cost_won", 800),
    ("monthly_plans", "representative_item", "수정 품목"),
    ("monthly_actual_quantities", "quantity_ea_text", "9"),
    ("monthly_actual_quantities", "source_note", "수정 출처"),
    ("monthly_sales", "amount_won", 900),
    ("monthly_sales", "confirmed_at", "2026-08-30T00:00:00+09:00"),
    ("monthly_sales", "source_note", "수정 매출 출처"),
    ("destinations", "name", "수정 납품처"),
    ("destinations", "display_order", 99),
    ("destinations", "active", 0),
    ("destinations", "required_for_report", 0),
    ("destinations", "representative_item", "새 기본 품목"),
    ("destinations", "include_quantity_total", 0),
    ("destinations", "include_cost_total", 0),
    ("destinations", "include_sales_total", 0),
])
def test_monthly_revision_covers_each_editable_value_and_master(client, masters, database, table, column, value):
    path = "/months/2026-08/inputs"
    assert post(client, path, monthly_form(masters)).status_code == 303
    before_revision = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    with database.connection() as connection:
        connection.execute(f"UPDATE {table} SET {column} = ? WHERE id = 1", (value,))
        connection.commit()
    after_revision = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    assert before_revision != after_revision
    assert re.fullmatch("[0-9a-f]{64}", after_revision)
    assert hidden_fields(client.get(path).text)["monthly_revision"] == after_revision
    before = monthly_database_state(database)
    response = post(client, path, monthly_form(masters) + [("monthly_revision", before_revision)])
    assert response.status_code == 409
    assert monthly_database_state(database) == before


def test_monthly_revision_distinguishes_blank_and_zero_and_is_required(client, masters, database):
    path = "/months/2026-08/inputs"
    blank_revision = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    assert post(client, path, monthly_form(masters)).status_code == 303
    zero_revision = hidden_fields(client.get(path).text).get("monthly_revision", "missing")
    assert blank_revision != zero_revision
    before = monthly_database_state(database)
    assert post(client, path, monthly_form(masters), revision=False).status_code == 409
    assert monthly_database_state(database) == before
