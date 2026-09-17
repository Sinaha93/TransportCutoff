ALTER TABLE transport_entries
ADD COLUMN source_alias TEXT CHECK(
  source_alias IS NULL OR (
    typeof(source_alias) = 'text' AND length(trim(source_alias)) > 0
  )
);

UPDATE transport_entries
SET source_alias = unresolved_alias
WHERE unresolved_alias IS NOT NULL AND length(trim(unresolved_alias)) > 0;
