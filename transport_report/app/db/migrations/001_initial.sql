CREATE TABLE schema_migrations(
  version INTEGER PRIMARY KEY,
  applied_at TEXT NOT NULL
);

CREATE TABLE destinations(
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL,
  representative_item TEXT,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0, 1)),
  required_for_report INTEGER NOT NULL DEFAULT 1 CHECK(required_for_report IN (0, 1)),
  include_quantity_total INTEGER NOT NULL DEFAULT 1 CHECK(include_quantity_total IN (0, 1)),
  include_cost_total INTEGER NOT NULL DEFAULT 1 CHECK(include_cost_total IN (0, 1)),
  include_sales_total INTEGER NOT NULL DEFAULT 1 CHECK(include_sales_total IN (0, 1))
);

CREATE TABLE destination_aliases(
  id INTEGER PRIMARY KEY,
  raw_name TEXT NOT NULL,
  source_type TEXT NOT NULL,
  destination_id INTEGER NOT NULL REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  UNIQUE(raw_name, source_type)
);

CREATE TABLE vehicle_rates(
  id INTEGER PRIMARY KEY,
  destination_id INTEGER NOT NULL REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  vehicle_type TEXT NOT NULL,
  unit_rate_won INTEGER NOT NULL CHECK(
    typeof(unit_rate_won) = 'integer' AND unit_rate_won >= 0
  ),
  effective_from TEXT NOT NULL CHECK(
    length(effective_from) = 7
    AND substr(effective_from, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(effective_from, 5, 1) = '-'
    AND substr(effective_from, 6, 2) BETWEEN '01' AND '12'
  ),
  effective_to TEXT CHECK(
    effective_to IS NULL OR (
      length(effective_to) = 7
      AND substr(effective_to, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
      AND substr(effective_to, 5, 1) = '-'
      AND substr(effective_to, 6, 2) BETWEEN '01' AND '12'
    )
  ),
  source_note TEXT,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0, 1)),
  UNIQUE(destination_id, vehicle_type, effective_from),
  CHECK(effective_to IS NULL OR effective_to >= effective_from)
);

CREATE TABLE report_groups(
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0, 1))
);

CREATE TABLE report_group_members(
  group_id INTEGER NOT NULL REFERENCES report_groups(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  destination_id INTEGER NOT NULL REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  include_quantity INTEGER NOT NULL DEFAULT 1 CHECK(include_quantity IN (0, 1)),
  include_cost INTEGER NOT NULL DEFAULT 1 CHECK(include_cost IN (0, 1)),
  PRIMARY KEY(group_id, destination_id)
);

CREATE TABLE monthly_plans(
  id INTEGER PRIMARY KEY,
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER NOT NULL REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  quantity_ea_text TEXT NOT NULL CHECK(
    typeof(quantity_ea_text) = 'text'
    AND quantity_ea_text = trim(quantity_ea_text)
    AND quantity_ea_text GLOB '*[0-9]*'
    AND quantity_ea_text NOT GLOB '*[^0-9.]*'
    AND length(quantity_ea_text) - length(replace(quantity_ea_text, '.', '')) <= 1
  ),
  cost_won INTEGER NOT NULL CHECK(typeof(cost_won) = 'integer' AND cost_won >= 0),
  representative_item TEXT,
  source_note TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, destination_id)
);

CREATE TABLE monthly_actual_quantities(
  id INTEGER PRIMARY KEY,
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER NOT NULL REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  quantity_ea_text TEXT NOT NULL CHECK(
    typeof(quantity_ea_text) = 'text'
    AND quantity_ea_text = trim(quantity_ea_text)
    AND quantity_ea_text GLOB '*[0-9]*'
    AND quantity_ea_text NOT GLOB '*[^0-9.]*'
    AND length(quantity_ea_text) - length(replace(quantity_ea_text, '.', '')) <= 1
  ),
  source_note TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, destination_id)
);

CREATE TABLE monthly_sales(
  id INTEGER PRIMARY KEY,
  report_month TEXT NOT NULL UNIQUE CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  amount_won INTEGER NOT NULL CHECK(
    typeof(amount_won) = 'integer' AND amount_won >= 0
  ),
  source_note TEXT,
  confirmed_at TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE import_batches(
  id INTEGER PRIMARY KEY,
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  source_type TEXT NOT NULL,
  source_filename TEXT NOT NULL,
  file_sha256 TEXT NOT NULL CHECK(length(file_sha256) = 64),
  status TEXT NOT NULL,
  error_summary TEXT,
  imported_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, source_type, file_sha256)
);

CREATE TABLE transport_entries(
  id INTEGER PRIMARY KEY,
  import_batch_id INTEGER NOT NULL REFERENCES import_batches(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  unresolved_alias TEXT,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL CHECK(source_row > 0),
  transport_day INTEGER NOT NULL CHECK(transport_day BETWEEN 1 AND 31),
  transport_type TEXT NOT NULL,
  vehicle_type TEXT,
  vehicle_driver_group TEXT,
  trip_count_text TEXT NOT NULL CHECK(
    typeof(trip_count_text) = 'text'
    AND trip_count_text = trim(trip_count_text)
    AND trip_count_text GLOB '*[0-9]*'
    AND trip_count_text NOT GLOB '*[^0-9.]*'
    AND length(trip_count_text) - length(replace(trip_count_text, '.', '')) <= 1
  ),
  unit_rate_won INTEGER CHECK(
    unit_rate_won IS NULL OR (
      typeof(unit_rate_won) = 'integer' AND unit_rate_won >= 0
    )
  ),
  cost_won INTEGER NOT NULL CHECK(typeof(cost_won) = 'integer' AND cost_won >= 0),
  source_note TEXT,
  UNIQUE(import_batch_id, source_sheet, source_row, transport_day, transport_type),
  CHECK(
    destination_id IS NOT NULL OR (
      unresolved_alias IS NOT NULL AND length(trim(unresolved_alias)) > 0
    )
  )
);

CREATE TABLE report_runs(
  id INTEGER PRIMARY KEY,
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  status TEXT NOT NULL,
  is_locked INTEGER NOT NULL DEFAULT 0 CHECK(is_locked IN (0, 1)),
  locked_at TEXT,
  unlocked_at TEXT,
  unlock_reason TEXT,
  input_revision TEXT,
  output_hashes_json TEXT,
  error_summary TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  CHECK(is_locked = 0 OR locked_at IS NOT NULL),
  CHECK(
    unlocked_at IS NULL OR (
      unlock_reason IS NOT NULL AND length(trim(unlock_reason)) > 0
    )
  )
);

CREATE INDEX idx_transport_entries_month_destination
  ON transport_entries(report_month, destination_id);
CREATE INDEX idx_report_runs_month_created
  ON report_runs(report_month, created_at);
