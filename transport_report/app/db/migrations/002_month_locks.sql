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
  input_revision INTEGER NOT NULL CHECK(
    typeof(input_revision) = 'integer'
    AND input_revision BETWEEN 0 AND 9223372036854775807
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

CREATE TABLE month_lock_migration_guard(value INTEGER);

CREATE TRIGGER month_lock_migration_reject_invalid_revision
BEFORE INSERT ON month_lock_migration_guard
WHEN NEW.value = 0
BEGIN
  SELECT RAISE(ABORT, 'invalid legacy lock input_revision');
END;

INSERT INTO month_lock_migration_guard(value)
SELECT CASE WHEN EXISTS (
  SELECT 1
  FROM report_runs AS candidate
  WHERE candidate.is_locked = 1
    AND candidate.id = (
      SELECT MAX(latest.id)
      FROM report_runs AS latest
      WHERE latest.report_month = candidate.report_month
        AND latest.is_locked = 1
    )
    AND (
      candidate.input_revision IS NULL
      OR length(trim(candidate.input_revision)) = 0
      OR trim(candidate.input_revision) GLOB '*[^0-9]*'
      OR length(ltrim(trim(candidate.input_revision), '0')) > 19
      OR (
        length(ltrim(trim(candidate.input_revision), '0')) = 19
        AND ltrim(trim(candidate.input_revision), '0') > '9223372036854775807'
      )
    )
) THEN 0 ELSE 1 END;

DROP TRIGGER month_lock_migration_reject_invalid_revision;
DROP TABLE month_lock_migration_guard;

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
  CAST(trim(legacy.input_revision) AS INTEGER),
  NULL
FROM report_runs AS legacy
WHERE legacy.is_locked = 1
  AND legacy.id = (
    SELECT MAX(latest.id)
    FROM report_runs AS latest
    WHERE latest.report_month = legacy.report_month
      AND latest.is_locked = 1
  );

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

CREATE TRIGGER report_runs_reject_month_lock_audit_delete
BEFORE DELETE ON report_runs
WHEN OLD.status IN ('month_locked', 'month_unlocked')
BEGIN
  SELECT RAISE(ABORT, 'month lock audit immutable');
END;
