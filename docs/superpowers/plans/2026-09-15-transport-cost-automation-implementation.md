# Transport Cost Reporting Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use $superpower-subagents (recommended) or $superpower-executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking via update_plan.

**Goal:** Build a portable Windows local web application that imports transport-cost workbooks, accepts monthly plan, quantity, and sales inputs, calculates the monthly report, and produces the existing six-slide PowerPoint format plus a review Excel file.

**Architecture:** A FastAPI application bound only to `127.0.0.1` serves server-rendered management screens backed by a portable SQLite database. Domain services normalize uploaded Excel data, calculate the monthly report, validate prerequisites, render three charts, and update a copy of the existing PowerPoint template without changing slide layout. A PyInstaller one-folder bundle packages Python and all dependencies so the application can be copied to another Windows PC and run without installation.

**Tech Stack:** Python, FastAPI, Uvicorn, Jinja2, SQLite, openpyxl, python-pptx, matplotlib, Pillow, PyInstaller, pytest.

---

## File structure

```text
transport_report/
├─ app/
│  ├─ main.py                         # FastAPI application and route registration
│  ├─ launcher.py                     # portable executable entry point
│  ├─ config.py                       # relative runtime paths and settings
│  ├─ db.py                           # SQLite connection and migrations
│  ├─ domain/
│  │  ├─ models.py                    # immutable domain records and report view models
│  │  ├─ calculations.py              # report calculations and aggregates
│  │  └─ validation.py                # blocking and warning validation rules
│  ├─ repositories/
│  │  ├─ masters.py                   # destinations, aliases, rates, groups
│  │  ├─ monthly_inputs.py            # plans, manual quantities, sales
│  │  └─ transport_entries.py         # normalized workbook import rows
│  ├─ importers/
│  │  ├─ protocols.py                 # transport and quantity source interfaces
│  │  ├─ hwaseong_workbook.py         # current transport workbook parser
│  │  └─ legacy_migration.py          # one-time migration from current planning workbook
│  ├─ reporting/
│  │  ├─ charts.py                    # three PPT chart images
│  │  ├─ pptx_report.py               # six-slide template updater
│  │  └─ xlsx_review.py               # review workbook exporter
│  ├─ services/
│  │  ├─ import_service.py            # hash, deduplication, staging, commit
│  │  ├─ report_service.py            # calculate, validate, and generate outputs
│  │  └─ backup_service.py            # SQLite backup and restore
│  ├─ web/
│  │  ├─ routes.py                    # browser page and form routes
│  │  ├─ templates/                   # Jinja2 pages
│  │  └─ static/                      # local CSS and JavaScript
│  └─ db/migrations/
│     └─ 001_initial.sql              # initial schema
├─ assets/
│  └─ report_template.pptx            # user-supplied six-slide template
├─ runtime/
│  ├─ data/app.db                     # created on first run
│  ├─ backups/                        # user-created database backups
│  ├─ imports/                        # copied source files by hash
│  └─ outputs/                        # generated XLSX, PPTX, and chart images
├─ tests/
│  ├─ fixtures/                       # synthetic workbooks and PPT template
│  ├─ unit/
│  └─ integration/
├─ tools/
│  ├─ build_portable.ps1              # PyInstaller build
│  └─ verify_ppt_layout.py            # shape and position comparison
├─ pyproject.toml
├─ requirements.lock
└─ README.md
```

## Task 1: Project scaffold and portable path policy

**Files:**
- Create: `transport_report/pyproject.toml`
- Create: `transport_report/app/config.py`
- Create: `transport_report/app/main.py`
- Create: `transport_report/tests/unit/test_config.py`
- Create: `.gitignore`

- [x] **Step 1: Protect business files and runtime data from version control**

Create `.gitignore` with:

```gitignore
*.xlsx
*.xls
*.pptx
transport_report/runtime/
transport_report/dist/
transport_report/build/
transport_report/.pytest_cache/
transport_report/**/__pycache__/
transport_report/.venv/
```

- [x] **Step 2: Initialize local version control**

Run:

```powershell
git init
git status --short
```

Expected: an empty repository with the existing business workbooks ignored after `.gitignore` is added.

- [x] **Step 3: Write the failing portable-path test**

```python
from pathlib import Path
from app.config import RuntimePaths


def test_runtime_paths_are_relative_to_bundle(tmp_path: Path):
    paths = RuntimePaths.from_root(tmp_path)
    assert paths.database == tmp_path / "runtime" / "data" / "app.db"
    assert paths.outputs == tmp_path / "runtime" / "outputs"
    assert paths.template == tmp_path / "assets" / "report_template.pptx"
```

- [x] **Step 4: Run the test and verify failure**

Run: `python -m pytest tests/unit/test_config.py -v`

Expected: FAIL because `app.config` does not exist.

- [x] **Step 5: Implement the path object and minimal app**

```python
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RuntimePaths:
    root: Path
    database: Path
    outputs: Path
    imports: Path
    backups: Path
    template: Path

    @classmethod
    def from_root(cls, root: Path) -> "RuntimePaths":
        return cls(
            root=root,
            database=root / "runtime" / "data" / "app.db",
            outputs=root / "runtime" / "outputs",
            imports=root / "runtime" / "imports",
            backups=root / "runtime" / "backups",
            template=root / "assets" / "report_template.pptx",
        )

    def ensure(self) -> None:
        for path in (self.database.parent, self.outputs, self.imports, self.backups):
            path.mkdir(parents=True, exist_ok=True)
```

Create `app/main.py` with a FastAPI health route returning `{"status": "ok"}`.

- [x] **Step 6: Run the focused test**

Run: `python -m pytest tests/unit/test_config.py -v`

Expected: PASS.

- [x] **Step 7: Commit the scaffold**

```powershell
git add .gitignore transport_report/pyproject.toml transport_report/app transport_report/tests/unit/test_config.py
git commit -m "chore: scaffold portable reporting app"
```

## Task 2: SQLite schema and migration runner

**Files:**
- Create: `transport_report/app/db.py`
- Create: `transport_report/app/db/migrations/001_initial.sql`
- Create: `transport_report/tests/unit/test_db.py`

- [x] **Step 1: Write the failing migration test**

```python
def test_initial_migration_creates_required_tables(tmp_path):
    db = Database(tmp_path / "app.db")
    db.migrate()
    names = db.table_names()
    assert {
        "destinations", "destination_aliases", "vehicle_rates",
        "report_groups", "report_group_members", "monthly_plans",
        "monthly_actual_quantities", "monthly_sales",
        "transport_entries", "import_batches", "report_runs",
    } <= names
```

- [x] **Step 2: Run the test and verify failure**

Run: `python -m pytest tests/unit/test_db.py::test_initial_migration_creates_required_tables -v`

Expected: FAIL because `Database` is missing.

- [x] **Step 3: Create the versioned schema**

The migration must define:

```sql
CREATE TABLE schema_migrations(version INTEGER PRIMARY KEY, applied_at TEXT NOT NULL);
CREATE TABLE destinations(
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL,
  active INTEGER NOT NULL DEFAULT 1,
  required_for_report INTEGER NOT NULL DEFAULT 1,
  include_quantity_total INTEGER NOT NULL DEFAULT 1,
  include_cost_total INTEGER NOT NULL DEFAULT 1,
  include_sales_total INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE destination_aliases(
  id INTEGER PRIMARY KEY,
  raw_name TEXT NOT NULL,
  source_type TEXT NOT NULL,
  destination_id INTEGER NOT NULL REFERENCES destinations(id),
  UNIQUE(raw_name, source_type)
);
CREATE TABLE report_groups(
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL,
  active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE report_group_members(
  group_id INTEGER NOT NULL REFERENCES report_groups(id),
  destination_id INTEGER NOT NULL REFERENCES destinations(id),
  include_quantity INTEGER NOT NULL DEFAULT 1,
  include_cost INTEGER NOT NULL DEFAULT 1,
  PRIMARY KEY(group_id, destination_id)
);
```

Also create the monthly, transport, import, and report-run tables with `report_month` stored as `YYYY-MM`, monetary values stored as integer won, quantities stored as decimal text, and unique keys preventing duplicate month/destination records.

- [x] **Step 4: Implement transactional migrations**

`Database.migrate()` must load migration files in numeric order, apply each migration inside one transaction, and record the version only after success.

- [x] **Step 5: Test repeatability**

Add and run a test that calls `migrate()` twice and confirms the second call makes no schema changes.

Run: `python -m pytest tests/unit/test_db.py -v`

Expected: PASS.

- [x] **Step 6: Commit the database foundation**

```powershell
git add transport_report/app/db.py transport_report/app/db/migrations transport_report/tests/unit/test_db.py
git commit -m "feat: add portable sqlite schema"
```

## Task 3: Domain models and master-data repositories

**Files:**
- Create: `transport_report/app/domain/models.py`
- Create: `transport_report/app/repositories/masters.py`
- Create: `transport_report/tests/unit/test_masters.py`

- [x] **Step 1: Write failing alias and group tests**

```python
def test_alias_resolves_to_standard_destination(repo):
    destination = repo.create_destination("포레시아 영천", display_order=10)
    repo.add_alias(destination.id, "포레시아-영천", "transport_workbook")
    assert repo.resolve_alias("포레시아-영천", "transport_workbook") == destination


def test_group_members_can_be_changed_without_formula_edits(repo):
    group = repo.create_group("영남권", display_order=20)
    members = [repo.create_destination(name, i) for i, name in enumerate(
        ["현대 울산", "포레시아 영천", "세종공업"], start=1
    )]
    repo.replace_group_members(group.id, [m.id for m in members])
    assert [m.name for m in repo.group_members(group.id)] == [
        "현대 울산", "포레시아 영천", "세종공업"
    ]
```

- [x] **Step 2: Run the tests and verify failure**

Run: `python -m pytest tests/unit/test_masters.py -v`

Expected: FAIL because the repository is missing.

- [x] **Step 3: Implement destination, alias, rate, and group repositories**

Use parameterized SQLite statements. Normalize aliases by trimming whitespace and Unicode-normalizing with NFKC, but do not remove meaningful punctuation automatically. Exact alias decisions remain user-controlled.

- [x] **Step 4: Seed the known group and exception**

Provide an idempotent seed function that creates:

```python
DEFAULT_GROUPS = {
    "영남권": ["현대 울산", "포레시아 영천", "세종공업"],
}
DEFAULT_TOTAL_RULES = {
    "당진": {"quantity": False, "cost": True, "sales": False},
}
```

- [x] **Step 5: Run repository tests**

Run: `python -m pytest tests/unit/test_masters.py -v`

Expected: PASS, including cascade-protection tests that prevent deleting a destination already used by monthly data.

- [x] **Step 6: Commit master-data support**

```powershell
git add transport_report/app/domain transport_report/app/repositories/masters.py transport_report/tests/unit/test_masters.py
git commit -m "feat: add editable destination and group masters"
```

## Task 4: Manual monthly inputs and future ERP extension boundary

**Files:**
- Create: `transport_report/app/importers/protocols.py`
- Create: `transport_report/app/repositories/monthly_inputs.py`
- Create: `transport_report/tests/unit/test_monthly_inputs.py`

- [x] **Step 1: Define and test the quantity-source contract**

```python
class QuantityProvider(Protocol):
    def load(self, report_month: str) -> list[QuantityRecord]: ...


@dataclass(frozen=True)
class QuantityRecord:
    report_month: str
    destination_id: int
    quantity_ea: Decimal
    source: str
```

Write a test proving `ManualQuantityProvider` returns explicitly entered zero while leaving an unentered destination absent.

- [x] **Step 2: Run the test and verify failure**

Run: `python -m pytest tests/unit/test_monthly_inputs.py -v`

Expected: FAIL because the provider and repository are missing.

- [x] **Step 3: Implement monthly input upserts**

Implement separate upserts for:

```python
save_plan(report_month, destination_id, quantity_ea, cost_won, representative_item)
save_actual_quantity(report_month, destination_id, quantity_ea, source_note="ERP 수기 확인")
save_sales(report_month, amount_won, source_note, confirmed_at)
```

Use `None` for missing quantity and `Decimal("0")` for confirmed zero.

- [x] **Step 4: Add month-lock behavior**

A finalized report month may be unlocked deliberately, but ordinary form submissions must reject updates while locked. Record the unlock and reason in `report_runs`.

- [x] **Step 5: Run focused tests**

Run: `python -m pytest tests/unit/test_monthly_inputs.py -v`

Expected: PASS for insert, update, explicit zero, missing value, and locked-month behavior.

- [x] **Step 6: Commit monthly inputs**

```powershell
git add transport_report/app/importers/protocols.py transport_report/app/repositories/monthly_inputs.py transport_report/tests/unit/test_monthly_inputs.py
git commit -m "feat: add manual monthly quantity and sales inputs"
```

## Task 5: Transport workbook importer

**Files:**
- Create: `transport_report/app/importers/hwaseong_workbook.py`
- Create: `transport_report/app/repositories/transport_entries.py`
- Create: `transport_report/app/services/import_service.py`
- Create: `transport_report/tests/fixtures/build_transport_fixture.py`
- Create: `transport_report/tests/unit/test_transport_importer.py`

- [x] **Step 1: Build a synthetic workbook fixture**

Generate a workbook containing:

- `화성운반비내역(8월)` with day columns `C:AG`, total count `AH`, unit cost `AJ`, subtotal `AK`.
- Repeated hidden header rows to verify they are skipped.
- `용차1-OK로지웰`, `용차2-대원로지스틱`, and `용차3-정동물류` with date, destination, tonnage, and amount.
- One known alias and one unknown alias.
- Fractional trip count `1.25`.

- [x] **Step 2: Write failing parser tests**

```python
def test_parser_normalizes_day_columns_and_fractional_trips(fixture_path):
    rows = HwaseongWorkbookParser().parse(fixture_path, "2026-08")
    assert any(r.day == 12 and r.trip_count == Decimal("1.25") for r in rows)


def test_import_is_idempotent(import_service, fixture_path):
    first = import_service.import_transport(fixture_path, "2026-08")
    second = import_service.import_transport(fixture_path, "2026-08")
    assert first.inserted_count > 0
    assert second.status == "duplicate"
```

- [x] **Step 3: Run the tests and verify failure**

Run: `python -m pytest tests/unit/test_transport_importer.py -v`

Expected: FAIL because the parser is missing.

- [x] **Step 4: Implement safe workbook parsing**

Open uploaded workbooks with `openpyxl.load_workbook(..., data_only=True, read_only=False, keep_vba=False)`. Never execute macros or external links. Reject files missing required sheets or headers.

For the regular sheet:

- scan rows with a destination label in column `B` and numeric unit cost in `AJ`;
- expand nonzero day cells `C:AG` into normalized entries;
- carry the latest nonempty vehicle/driver group from column `A` as descriptive metadata;
- calculate `cost_won = trip_count × unit_cost_won`;
- compare the calculated total to cached `AK` and raise a blocking mismatch if they differ.

For each subcontracted-car sheet, read detail rows only and exclude displayed total rows. Use amount column `I` as nonregular transport cost.

- [x] **Step 5: Implement staging and alias resolution**

Store the SHA-256 file hash in `import_batches`. Stage all rows first, resolve aliases, and commit the batch only when structural checks pass. Unknown aliases remain imported with `destination_id=NULL` and become blocking validation errors.

- [x] **Step 6: Run parser tests**

Run: `python -m pytest tests/unit/test_transport_importer.py -v`

Expected: PASS for normal rows, fractional trips, repeated headers, unknown aliases, duplicate imports, and subtotal mismatch.

- [x] **Step 7: Commit the importer**

```powershell
git add transport_report/app/importers/hwaseong_workbook.py transport_report/app/repositories/transport_entries.py transport_report/app/services/import_service.py transport_report/tests
git commit -m "feat: import and normalize transport workbooks"
```

## Task 6: Monthly calculation engine

**Files:**
- Create: `transport_report/app/domain/calculations.py`
- Create: `transport_report/tests/unit/test_calculations.py`

- [x] **Step 1: Write failing destination calculations**

```python
def test_unit_transport_cost_is_cost_divided_by_quantity():
    result = calculate_destination(
        planned_quantity=Decimal("15152"),
        planned_cost_won=1_578_333,
        actual_quantity=Decimal("14782"),
        actual_cost_won=1_766_000,
    )
    assert result.planned_unit_cost == Decimal("104.17")
    assert result.actual_unit_cost == Decimal("119.47")


def test_zero_quantity_produces_no_unit_cost_instead_of_division_error():
    result = calculate_destination(Decimal("46"), 254_000, Decimal("0"), 0)
    assert result.actual_unit_cost is None
    assert result.actual_unit_cost_variance_pct is None
```

- [x] **Step 2: Write failing total and group tests**

Test that 당진 is excluded from quantity totals but included in cost totals. Test that 영남권 sums 현대 울산, 포레시아 영천, and 세종공업 without being added again to the grand total.

- [x] **Step 3: Run the tests and verify failure**

Run: `python -m pytest tests/unit/test_calculations.py -v`

Expected: FAIL because the calculation functions are missing.

- [x] **Step 4: Implement calculation functions**

Use `Decimal` throughout and centralize rounding:

```python
def unit_cost(cost_won: int, quantity_ea: Decimal) -> Decimal | None:
    if quantity_ea == 0:
        return None
    return (Decimal(cost_won) / quantity_ea).quantize(Decimal("0.01"))


def variance(actual: Decimal, plan: Decimal) -> Decimal:
    return actual - plan


def variance_pct(actual: Decimal, plan: Decimal) -> Decimal | None:
    return None if plan == 0 else (actual - plan) / plan
```

Calculate 3-, 6-, and 12-month averages from complete calendar months ending one month before the configured report month. For a 2026-08 report, use 2026-05 through 2026-07 for the 3-month average, 2026-02 through 2026-07 for the 6-month average, and 2025-08 through 2026-07 for the 12-month average. Calculate the comparison-year average from the full year immediately preceding the report year.

- [x] **Step 5: Add the ±15% review selector**

Select destinations where unit-cost variance percent is less than or equal to `-0.15` or greater than or equal to `0.15`. A missing unit cost becomes a validation item, not an automatic narrative.

- [x] **Step 6: Run calculation tests**

Run: `python -m pytest tests/unit/test_calculations.py -v`

Expected: PASS for destination, total, group, average, zero, and threshold cases.

- [x] **Step 7: Commit calculations**

```powershell
git add transport_report/app/domain/calculations.py transport_report/tests/unit/test_calculations.py
git commit -m "feat: calculate monthly transport report"
```

## Task 7: Blocking validation engine

**Files:**
- Create: `transport_report/app/domain/validation.py`
- Create: `transport_report/tests/unit/test_validation.py`

- [x] **Step 1: Write one failing test per blocking rule**

Cover these exact codes:

```python
MISSING_DESTINATION_ALIAS
MISSING_VEHICLE_RATE
MISSING_ACTUAL_QUANTITY
MISSING_SALES
DUPLICATE_IMPORT
TRANSPORT_SUBTOTAL_MISMATCH
TOTAL_RECONCILIATION_FAILED
MISSING_NEXT_MONTH_PLAN
MISSING_PRIOR_YEAR_HISTORY
REPORT_MONTH_MISMATCH
```

Assert that explicit quantity zero does not produce `MISSING_ACTUAL_QUANTITY`.

- [x] **Step 2: Run the tests and verify failure**

Run: `python -m pytest tests/unit/test_validation.py -v`

Expected: FAIL because `validate_report()` is missing.

- [x] **Step 3: Implement structured validation results**

```python
@dataclass(frozen=True)
class ValidationIssue:
    code: str
    severity: Literal["error", "warning"]
    message: str
    destination_id: int | None = None
    source_locator: str | None = None
```

`validate_report()` returns issues sorted by severity, destination display order, and code. `can_generate` is true only when no error-severity issues exist.

- [x] **Step 4: Run validation tests**

Run: `python -m pytest tests/unit/test_validation.py -v`

Expected: PASS with stable error codes and user-readable Korean messages.

- [x] **Step 5: Commit validation**

```powershell
git add transport_report/app/domain/validation.py transport_report/tests/unit/test_validation.py
git commit -m "feat: block invalid report generation"
```

## Task 8: Legacy data migration

**Files:**
- Create: `transport_report/app/importers/legacy_migration.py`
- Create: `transport_report/tests/unit/test_legacy_migration.py`

- [x] **Step 1: Write a synthetic legacy fixture and failing migration test**

The fixture must include `누적 데이터` columns `A:F` for year, month, destination, actual quantity, and actual cost; `26년 월계획` month-pair columns; and the current 8월 destination values.

Assert migration produces historical monthly rows, current plans, actual quantities, and costs without importing chart helper ranges.

- [x] **Step 2: Run the test and verify failure**

Run: `python -m pytest tests/unit/test_legacy_migration.py -v`

Expected: FAIL because the importer is missing.

- [x] **Step 3: Implement migration from current files**

Read `누적 데이터!A:F` as historical destination-month actuals and `26년 월계획` as plan pairs. Import only values, not legacy formulas. Produce a reconciliation report comparing:

- August planned quantity and cost;
- August actual quantity and cost;
- grand totals excluding/including destination rules;
- 3-, 6-, and 12-month averages;
- prior-year averages.

- [x] **Step 4: Add dry-run and commit modes**

Dry run returns counts, unknown aliases, and reconciliation differences without changing SQLite. Commit mode writes only after the user confirms the dry-run result in the UI.

- [x] **Step 5: Run migration tests**

Run: `python -m pytest tests/unit/test_legacy_migration.py -v`

Expected: PASS for dry run, commit, rollback on mismatch, and repeated migration.

- [x] **Step 6: Commit migration support**

```powershell
git add transport_report/app/importers/legacy_migration.py transport_report/tests/unit/test_legacy_migration.py
git commit -m "feat: migrate historical transport data"
```

## Task 9: Local management web UI

**Files:**
- Create: `transport_report/app/web/routes.py`
- Create: `transport_report/app/web/templates/base.html`
- Create: `transport_report/app/web/templates/dashboard.html`
- Create: `transport_report/app/web/templates/destinations.html`
- Create: `transport_report/app/web/templates/groups.html`
- Create: `transport_report/app/web/templates/monthly_inputs.html`
- Create: `transport_report/app/web/templates/import_review.html`
- Create: `transport_report/app/web/templates/report_preview.html`
- Create: `transport_report/app/web/static/app.css`
- Create: `transport_report/tests/integration/test_web_routes.py`

- [x] **Step 1: Write failing browser-route tests**

```python
def test_monthly_input_page_distinguishes_blank_and_zero(client):
    response = client.get("/months/2026-08/inputs")
    assert response.status_code == 200
    assert 'name="actual_quantity"' in response.text
    assert 'data-zero-is-valid="true"' in response.text


def test_unknown_aliases_are_actionable(client, seeded_unknown_alias):
    response = client.get("/imports/review")
    assert "미등록 명칭" in response.text
    assert "기존 납품처에 연결" in response.text
    assert "새 납품처 등록" in response.text
```

- [x] **Step 2: Run route tests and verify failure**

Run: `python -m pytest tests/integration/test_web_routes.py -v`

Expected: FAIL because routes and templates are missing.

- [x] **Step 3: Implement the dashboard workflow**

The first screen must show, in order:

1. report month selector;
2. import status;
3. plan, actual quantity, and sales completion status;
4. blocking validation issues;
5. `Excel 검토파일 생성` and `PPT 생성` actions.

Disable generation buttons server-side and visually when `can_generate` is false.

- [x] **Step 4: Implement master-management screens**

Destination screen supports name, aliases, display order, active state, representative item, and total inclusion flags. Group screen supports group name, ordered members, and per-member quantity/cost inclusion. Changes redirect to a recalculated report preview.

- [x] **Step 5: Implement monthly manual-entry screens**

Use one row per required destination with plan quantity, plan cost, representative item, actual quantity, and a source note. Put report-month sales and confirmation date above the table. Save all rows in one transaction.

- [x] **Step 6: Run web tests**

Run: `python -m pytest tests/integration/test_web_routes.py -v`

Expected: PASS for form validation, alias resolution, group edits, missing inputs, explicit zero, and generation-button state.

- [x] **Step 7: Commit the UI**

```powershell
git add transport_report/app/web transport_report/tests/integration/test_web_routes.py
git commit -m "feat: add local report management UI"
```

## Task 10: Chart rendering

**Files:**
- Create: `transport_report/app/reporting/charts.py`
- Create: `transport_report/tests/unit/test_charts.py`

- [ ] **Step 1: Write failing chart-data tests**

Assert that an August 2026 report produces month labels from `25년 8월` through `26년 8월`, followed by `3개월 평균`, `6개월 평균`, and `12개월 평균`, and uses `25년 평균` as the comparison line.

- [ ] **Step 2: Run the test and verify failure**

Run: `python -m pytest tests/unit/test_charts.py -v`

Expected: FAIL because chart builders are missing.

- [ ] **Step 3: Implement three deterministic PNG renderers**

Create:

```python
render_quantity_chart(report, path, dpi=200)
render_cost_chart(report, path, dpi=200)
render_combined_chart(report, path, dpi=200)
```

Use fixed fonts, colors, canvas sizes, legend order, axes, data labels, and bottom data tables matching the current presentation. Close every matplotlib figure after saving.

- [ ] **Step 4: Add image assertions**

Verify exact pixel dimensions, nonblank bounding boxes, expected legend labels, and that the final month values appear in the chart data model. Store no customer data in committed golden images; use synthetic fixtures.

- [ ] **Step 5: Run chart tests**

Run: `python -m pytest tests/unit/test_charts.py -v`

Expected: PASS and three PNGs generated in pytest temporary directories.

- [ ] **Step 6: Commit chart generation**

```powershell
git add transport_report/app/reporting/charts.py transport_report/tests/unit/test_charts.py
git commit -m "feat: render report charts"
```

## Task 11: PowerPoint template updater

**Files:**
- Create: `transport_report/app/reporting/pptx_report.py`
- Create: `transport_report/tools/verify_ppt_layout.py`
- Create: `transport_report/tests/integration/test_pptx_report.py`
- Copy during implementation: `26년 8월 운반비 보고.pptx` to `transport_report/assets/report_template.pptx`

- [ ] **Step 1: Create a sanitized test template**

Build a six-slide synthetic PPTX with the same shape roles and stable internal shape names but fictional values. Do not commit the business report as a test fixture.

- [ ] **Step 2: Write the failing generation test**

```python
def test_ppt_generation_updates_values_and_preserves_layout(template, report, tmp_path):
    output = tmp_path / "26년 8월 운반비 보고.pptx"
    generate_pptx(template, report, output)
    assert_presentation_text(output, "26년 8월 화성공장 운반비 보고")
    assert_presentation_text(output, "26년 9월 운반비 계획")
    assert_table_value(output, slide=2, row="합계", column="실적 운반비", value="27,355")
    assert_same_shape_geometry(template, output)
```

- [ ] **Step 3: Run the test and verify failure**

Run: `python -m pytest tests/integration/test_pptx_report.py -v`

Expected: FAIL because `generate_pptx()` is missing.

- [ ] **Step 4: Implement style-preserving text replacement**

Locate slides and shapes by slide number plus validated shape name. For text and table cells, replace run text while retaining the first run's font, size, bold, color, alignment, margins, and fill. Fail with `TEMPLATE_STRUCTURE_CHANGED` if a required shape is absent.

- [ ] **Step 5: Implement slide mappings**

- Slide 1: report title and creation date.
- Slide 2: monthly destination table, totals, averages, and selected ±15% review rows.
- Slides 3–5: remove only the named chart picture and insert the new PNG using the original picture's exact position, size, and crop.
- Slide 3: prior-year average quantity callout.
- Slide 4: prior-year average cost callout in thousand won.
- Slide 5: both callouts.
- Slide 6: next-month plan table and title derived from the next-month data, never copied from the report-month title.

- [ ] **Step 6: Verify layout programmatically**

`verify_ppt_layout.py` must compare slide count, required shape names, x/y/width/height, table row/column counts, and picture count against the template. Allow text values and image bytes to differ.

- [ ] **Step 7: Run PowerPoint tests**

Run: `python -m pytest tests/integration/test_pptx_report.py -v`

Expected: PASS, including report-month/next-month title checks and template-change failure.

- [ ] **Step 8: Commit PowerPoint generation**

```powershell
git add transport_report/app/reporting/pptx_report.py transport_report/tools/verify_ppt_layout.py transport_report/tests/integration/test_pptx_report.py
git commit -m "feat: generate existing six-slide report format"
```

## Task 12: Review Excel export

**Files:**
- Create: `transport_report/app/reporting/xlsx_review.py`
- Create: `transport_report/tests/integration/test_xlsx_review.py`

- [ ] **Step 1: Write the failing workbook test**

Assert the workbook contains `월간 종합`, `운행실적`, `검증 결과`, and `마스터 기준` sheets, with `월간 종합` first and formula-error strings absent.

- [ ] **Step 2: Run the test and verify failure**

Run: `python -m pytest tests/integration/test_xlsx_review.py -v`

Expected: FAIL because the exporter is missing.

- [ ] **Step 3: Implement the export**

Write calculated values as typed numbers, not string-formatted numbers. Include source file hash and import time in the input sheet. Freeze headers, apply filters, set sensible widths, and visually distinguish manual inputs, imported values, calculated results, warnings, and errors.

- [ ] **Step 4: Add reconciliation assertions**

Reload the saved workbook with `data_only=False` and assert totals equal the report model, dates remain dates, percentages remain numeric, and no `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` strings exist.

- [ ] **Step 5: Run Excel export tests**

Run: `python -m pytest tests/integration/test_xlsx_review.py -v`

Expected: PASS.

- [ ] **Step 6: Commit Excel export**

```powershell
git add transport_report/app/reporting/xlsx_review.py transport_report/tests/integration/test_xlsx_review.py
git commit -m "feat: export monthly review workbook"
```

## Task 13: Report orchestration, backup, and audit trail

**Files:**
- Create: `transport_report/app/services/report_service.py`
- Create: `transport_report/app/services/backup_service.py`
- Create: `transport_report/tests/integration/test_report_service.py`

- [ ] **Step 1: Write the failing end-to-end service tests**

Test that generation is blocked when sales, quantity, alias, prior-year history, or next-month plan is missing. Test that a valid month produces one XLSX, one PPTX, and three chart images under a timestamped output directory.

- [ ] **Step 2: Run the tests and verify failure**

Run: `python -m pytest tests/integration/test_report_service.py -v`

Expected: FAIL because report orchestration is missing.

- [ ] **Step 3: Implement report generation transaction**

The service must:

1. read a consistent database snapshot;
2. calculate the report;
3. validate all prerequisites;
4. stop without partial deliverables on errors;
5. create charts, XLSX, and PPTX in a temporary output directory;
6. verify generated files;
7. atomically rename the directory to `runtime/outputs/YYYY-MM/run-YYYYMMDD-HHMMSS`;
8. record file hashes and input revision in `report_runs`.

- [ ] **Step 4: Implement SQLite backup and guarded restore**

Use SQLite's online backup API. A backup file name contains the timestamp and schema version. Restore first validates the SQLite header and schema version, then creates an automatic pre-restore backup before replacing data.

- [ ] **Step 5: Run service tests**

Run: `python -m pytest tests/integration/test_report_service.py -v`

Expected: PASS for failure cleanup, successful output, audit hashes, backup, and restore.

- [ ] **Step 6: Commit orchestration**

```powershell
git add transport_report/app/services transport_report/tests/integration/test_report_service.py
git commit -m "feat: orchestrate validated report generation"
```

## Task 14: Portable launcher and Windows build

**Files:**
- Create: `transport_report/app/launcher.py`
- Create: `transport_report/tools/build_portable.ps1`
- Create: `transport_report/tests/unit/test_launcher.py`
- Create: `transport_report/README.md`

- [ ] **Step 1: Write failing launcher tests**

Test root discovery in source and PyInstaller modes, runtime directory creation, port selection, browser URL generation, and a second-launch result that opens the existing instance rather than starting another server.

- [ ] **Step 2: Run the tests and verify failure**

Run: `python -m pytest tests/unit/test_launcher.py -v`

Expected: FAIL because the launcher is missing.

- [ ] **Step 3: Implement a localhost-only launcher**

Bind to `127.0.0.1`, choose port `8765` or the next free port, store the active URL and process ID under `runtime`, open the default browser, and expose a UI shutdown action protected by a per-run random token. Never bind to `0.0.0.0`.

- [ ] **Step 4: Implement the PyInstaller one-folder build**

`build_portable.ps1` must:

```powershell
python -m pytest -q
python -m PyInstaller --noconfirm --clean --onedir --name TransportReport app/launcher.py
```

Then copy `web/templates`, `web/static`, database migrations, and `assets/report_template.pptx` into `dist/TransportReport`. Create empty `runtime/data`, `runtime/imports`, `runtime/backups`, and `runtime/outputs` directories.

- [ ] **Step 5: Document the user workflow**

README instructions must be limited to:

1. copy the complete `TransportReport` folder;
2. run `TransportReport.exe`;
3. allow Windows security review if company policy requires it;
4. manage masters and monthly inputs in the browser;
5. generate and open output files;
6. use the backup button before moving to another PC.

- [ ] **Step 6: Run launcher tests and build**

Run:

```powershell
python -m pytest tests/unit/test_launcher.py -v
powershell -ExecutionPolicy Bypass -File tools/build_portable.ps1
```

Expected: tests PASS and `dist/TransportReport/TransportReport.exe` exists.

- [ ] **Step 7: Commit packaging**

```powershell
git add transport_report/app/launcher.py transport_report/tools/build_portable.ps1 transport_report/tests/unit/test_launcher.py transport_report/README.md
git commit -m "build: package portable Windows application"
```

## Task 15: Real-file reconciliation and acceptance test

**Files:**
- Create: `transport_report/tests/manual/acceptance-checklist.md`
- Modify only if defects are found: files implemented in Tasks 1–14

- [ ] **Step 1: Make protected working copies of the supplied files**

Copy the three supplied Excel files and the PowerPoint into a temporary acceptance directory. Do not modify the originals.

- [ ] **Step 2: Run the legacy migration dry run**

Use the UI to import historical data from `26년 운반비 계획.xlsx`. Resolve aliases in the management screen and confirm the initial 영남권 group and 당진 inclusion rules.

Expected: no unresolved aliases and a reconciliation preview before database commit.

- [ ] **Step 3: Enter the confirmed manual inputs**

For August 2026, enter the ERP-confirmed quantities and the manually received sales amount. Import `화성공장 운행실적(26년8월).xlsx`. Enter the September plan.

- [ ] **Step 4: Reconcile the August monthly report**

Compare every row in the generated report against `8월 종합_수정!B6:T25`, including:

- each destination's plan and actual quantity;
- each destination's plan and actual transport cost;
- unit transport cost defined as cost divided by quantity;
- quantity, cost, and unit-cost variances;
- grand totals using the 당진 exception;
- sales and transport-cost-to-sales ratio;
- 3-, 6-, 12-month, and 25-year averages;
- 영남권 aggregation.

Expected: exact source values before presentation rounding and matching displayed values after rounding.

- [ ] **Step 5: Generate and inspect the PowerPoint**

Generate the August report and confirm:

- six slides and unchanged layout;
- slide 1 says August 2026;
- slide 2 values match the calculation model;
- slides 3–5 end at August 2026;
- slide 6 says September 2026 and matches September plan data;
- no `#DIV/0!` or stale month text is visible;
- Korean text and number labels are not clipped.

- [ ] **Step 6: Verify portability on the work PC**

Copy only `dist/TransportReport` to the user's work PC. Run the executable without Python, Node.js, database, or installer prerequisites. Generate the same report and compare SHA-256 hashes for the SQLite backup and calculation JSON; allow PPTX binary hashes to differ if package timestamps differ, but require equal extracted slide text, table values, image dimensions, and shape geometry.

- [ ] **Step 7: Run the complete automated suite**

Run: `python -m pytest -q`

Expected: all tests PASS with no skipped calculation, validation, report, or portability tests.

- [ ] **Step 8: Commit acceptance documentation and final fixes**

```powershell
git add transport_report/tests/manual/acceptance-checklist.md transport_report/app transport_report/tests transport_report/tools transport_report/README.md
git commit -m "test: verify August report and portable build"
```

## Future ERP quantity upgrade

The initial release ends with manual ERP quantity entry. A later upgrade adds these files without replacing the calculation or reporting layers:

```text
app/importers/erp_quantity_workbook.py
app/repositories/item_destination_mappings.py
web/templates/item_mappings.html
db/migrations/006_item_destination_mapping.sql
tests/unit/test_erp_quantity_importer.py
```

`ErpWorkbookQuantityProvider` will implement the existing `QuantityProvider` contract. It will parse item number and quantity from an ERP export, require every item number to map to one standard destination, aggregate quantities by destination and month, and present unmapped item numbers for user resolution before replacing manual quantities. Existing databases will be upgraded through migration `006` without losing master data, groups, reports, or history.

## Verification summary

- Unit tests cover paths, migrations, aliases, groups, manual inputs, imports, calculations, validation, chart data, launcher behavior, and explicit zero handling.
- Integration tests cover browser forms, legacy migration, Excel export, PowerPoint generation, report orchestration, backup, and restore.
- Real-file acceptance reconciles the August 2026 report cell by cell without editing original business files.
- Portable acceptance runs the packaged folder on the actual work PC without installing a runtime.
- PowerPoint verification checks both extracted content and unchanged slide geometry, followed by a manual visual review.
