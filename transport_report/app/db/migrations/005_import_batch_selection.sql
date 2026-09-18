ALTER TABLE import_batches
ADD COLUMN is_current INTEGER NOT NULL DEFAULT 0 CHECK(
  typeof(is_current) = 'integer' AND is_current IN (0, 1)
);

UPDATE import_batches
SET is_current = 1
WHERE status IN ('imported', 'imported_with_errors')
  AND id = (
    SELECT MAX(candidate.id)
    FROM import_batches AS candidate
    WHERE candidate.report_month = import_batches.report_month
      AND candidate.source_type = import_batches.source_type
      AND candidate.status IN ('imported', 'imported_with_errors')
  );

CREATE UNIQUE INDEX uq_import_batches_current_month_source
ON import_batches(report_month, source_type)
WHERE is_current = 1;
