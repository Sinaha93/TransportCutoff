ALTER TABLE transport_entries
ADD COLUMN source_alias TEXT CHECK(
  source_alias IS NULL OR (
    typeof(source_alias) = 'text' AND length(trim(source_alias)) > 0
  )
);

DROP TRIGGER transport_entries_reject_locked_update;

UPDATE transport_entries
SET source_alias = unresolved_alias
WHERE unresolved_alias IS NOT NULL AND length(trim(unresolved_alias)) > 0;

CREATE TRIGGER transport_entries_reject_locked_update
BEFORE UPDATE ON transport_entries
WHEN EXISTS (
  SELECT 1 FROM month_locks
  WHERE report_month IN (OLD.report_month, NEW.report_month) AND is_locked = 1
)
BEGIN
  SELECT RAISE(ABORT, 'report month locked');
END;
