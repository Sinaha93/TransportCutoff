CREATE TABLE month_locks(
  report_month TEXT PRIMARY KEY CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  is_locked INTEGER NOT NULL CHECK(
    typeof(is_locked) = 'integer' AND is_locked IN (0, 1)
  ),
  locked_at TEXT NOT NULL,
  unlocked_at TEXT,
  input_revision TEXT NOT NULL CHECK(
    typeof(input_revision) = 'text'
    AND length(trim(input_revision)) > 0
  ),
  last_unlock_reason TEXT,
  CHECK(
    (is_locked = 1 AND unlocked_at IS NULL)
    OR (
      is_locked = 0
      AND unlocked_at IS NOT NULL
      AND last_unlock_reason IS NOT NULL
      AND length(trim(last_unlock_reason)) > 0
    )
  )
);

INSERT INTO month_locks(
  report_month,
  is_locked,
  locked_at,
  unlocked_at,
  input_revision,
  last_unlock_reason
)
SELECT
  legacy.report_month,
  1,
  legacy.locked_at,
  NULL,
  legacy.input_revision,
  NULL
FROM report_runs AS legacy
WHERE legacy.is_locked = 1
  AND legacy.id = (
    SELECT MAX(latest.id)
    FROM report_runs AS latest
    WHERE latest.report_month = legacy.report_month
      AND latest.is_locked = 1
  );

INSERT INTO report_runs(
  report_month,
  status,
  is_locked,
  locked_at,
  input_revision
)
SELECT
  report_month,
  'month_locked',
  1,
  locked_at,
  input_revision
FROM month_locks;

CREATE TRIGGER monthly_plans_reject_locked_insert
BEFORE INSERT ON monthly_plans
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = NEW.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_plans_reject_locked_update
BEFORE UPDATE ON monthly_plans
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_plans_reject_locked_delete
BEFORE DELETE ON monthly_plans
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = OLD.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_quantities_reject_locked_insert
BEFORE INSERT ON monthly_actual_quantities
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = NEW.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_quantities_reject_locked_update
BEFORE UPDATE ON monthly_actual_quantities
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_quantities_reject_locked_delete
BEFORE DELETE ON monthly_actual_quantities
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = OLD.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_sales_reject_locked_insert
BEFORE INSERT ON monthly_sales
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = NEW.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_sales_reject_locked_update
BEFORE UPDATE ON monthly_sales
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_sales_reject_locked_delete
BEFORE DELETE ON monthly_sales
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = OLD.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_costs_reject_locked_insert
BEFORE INSERT ON monthly_actual_costs
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = NEW.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_costs_reject_locked_update
BEFORE UPDATE ON monthly_actual_costs
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER monthly_actual_costs_reject_locked_delete
BEFORE DELETE ON monthly_actual_costs
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = OLD.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER transport_entries_reject_locked_insert
BEFORE INSERT ON transport_entries
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = NEW.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER transport_entries_reject_locked_update
BEFORE UPDATE ON transport_entries
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER transport_entries_reject_locked_delete
BEFORE DELETE ON transport_entries
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month = OLD.report_month AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;

CREATE TRIGGER report_runs_reject_month_lock_audit_update
BEFORE UPDATE ON report_runs
WHEN OLD.status IN ('month_locked', 'month_unlocked')
  OR NEW.status IN ('month_locked', 'month_unlocked')
BEGIN
  SELECT RAISE(ABORT, 'month lock audit immutable');
END;

CREATE TRIGGER report_runs_reject_month_lock_audit_replace
BEFORE INSERT ON report_runs
WHEN EXISTS (
  SELECT 1
  FROM report_runs
  WHERE id = NEW.id AND status IN ('month_locked', 'month_unlocked')
)
BEGIN
  SELECT RAISE(ABORT, 'month lock audit immutable');
END;

CREATE TRIGGER report_runs_reject_month_lock_audit_delete
BEFORE DELETE ON report_runs
WHEN OLD.status IN ('month_locked', 'month_unlocked')
BEGIN
  SELECT RAISE(ABORT, 'month lock audit immutable');
END;
