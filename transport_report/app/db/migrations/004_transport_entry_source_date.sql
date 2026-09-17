ALTER TABLE transport_entries
ADD COLUMN source_date TEXT CHECK(
  source_date IS NULL OR (
    typeof(source_date) = 'text'
    AND length(source_date) = 10
    AND substr(source_date, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(source_date, 5, 1) = '-'
    AND substr(source_date, 6, 2) BETWEEN '01' AND '12'
    AND substr(source_date, 8, 1) = '-'
    AND substr(source_date, 9, 2) BETWEEN '01' AND '31'
    AND date(source_date, '+0 days') = source_date
  )
);
