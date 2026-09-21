CREATE TABLE report_supplemental_inputs(
  report_month TEXT PRIMARY KEY NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  planned_sales_won INTEGER CHECK(planned_sales_won IS NULL OR (typeof(planned_sales_won) = 'integer' AND planned_sales_won >= 0)),
  nonregular_planned_quantity_text TEXT CHECK(nonregular_planned_quantity_text IS NULL OR (
    typeof(nonregular_planned_quantity_text) = 'text' AND nonregular_planned_quantity_text NOT GLOB '*[^0-9.]*'
    AND length(nonregular_planned_quantity_text) - length(replace(nonregular_planned_quantity_text, '.', '')) <= 1
    AND (
      (instr(nonregular_planned_quantity_text, '.') = 0 AND (nonregular_planned_quantity_text = '0' OR substr(nonregular_planned_quantity_text, 1, 1) BETWEEN '1' AND '9'))
      OR (instr(nonregular_planned_quantity_text, '.') > 1
        AND (substr(nonregular_planned_quantity_text, 1, instr(nonregular_planned_quantity_text, '.') - 1) = '0' OR substr(nonregular_planned_quantity_text, 1, 1) BETWEEN '1' AND '9')
        AND substr(nonregular_planned_quantity_text, -1, 1) BETWEEN '1' AND '9')
    )
  )),
  nonregular_planned_cost_won INTEGER CHECK(nonregular_planned_cost_won IS NULL OR (typeof(nonregular_planned_cost_won) = 'integer' AND nonregular_planned_cost_won >= 0)),
  nonregular_actual_quantity_text TEXT CHECK(nonregular_actual_quantity_text IS NULL OR (
    typeof(nonregular_actual_quantity_text) = 'text' AND nonregular_actual_quantity_text NOT GLOB '*[^0-9.]*'
    AND length(nonregular_actual_quantity_text) - length(replace(nonregular_actual_quantity_text, '.', '')) <= 1
    AND (
      (instr(nonregular_actual_quantity_text, '.') = 0 AND (nonregular_actual_quantity_text = '0' OR substr(nonregular_actual_quantity_text, 1, 1) BETWEEN '1' AND '9'))
      OR (instr(nonregular_actual_quantity_text, '.') > 1
        AND (substr(nonregular_actual_quantity_text, 1, instr(nonregular_actual_quantity_text, '.') - 1) = '0' OR substr(nonregular_actual_quantity_text, 1, 1) BETWEEN '1' AND '9')
        AND substr(nonregular_actual_quantity_text, -1, 1) BETWEEN '1' AND '9')
    )
  )),
  nonregular_actual_cost_won INTEGER CHECK(nonregular_actual_cost_won IS NULL OR (typeof(nonregular_actual_cost_won) = 'integer' AND nonregular_actual_cost_won >= 0)),
  source_note TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  CHECK((nonregular_planned_quantity_text IS NULL) = (nonregular_planned_cost_won IS NULL)),
  CHECK((nonregular_actual_quantity_text IS NULL) = (nonregular_actual_cost_won IS NULL))
);

CREATE TRIGGER report_supplemental_inputs_reject_locked_insert
BEFORE INSERT ON report_supplemental_inputs
WHEN EXISTS (SELECT 1 FROM month_locks WHERE report_month IN (NEW.report_month) AND is_locked = 1)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER report_supplemental_inputs_reject_locked_update
BEFORE UPDATE ON report_supplemental_inputs
WHEN EXISTS (SELECT 1 FROM month_locks WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER report_supplemental_inputs_reject_locked_delete
BEFORE DELETE ON report_supplemental_inputs
WHEN EXISTS (SELECT 1 FROM month_locks WHERE report_month IN (OLD.report_month) AND is_locked = 1)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;
