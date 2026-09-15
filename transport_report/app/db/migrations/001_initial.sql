CREATE TABLE schema_migrations(
  version INTEGER PRIMARY KEY CHECK(typeof(version) = 'integer' AND version > 0),
  applied_at TEXT NOT NULL
);

CREATE TABLE destinations(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL CHECK(
    typeof(display_order) = 'integer' AND display_order >= 0
  ),
  representative_item TEXT,
  active INTEGER NOT NULL DEFAULT 1 CHECK(typeof(active) = 'integer' AND active IN (0, 1)),
  required_for_report INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(required_for_report) = 'integer' AND required_for_report IN (0, 1)
  ),
  include_quantity_total INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(include_quantity_total) = 'integer' AND include_quantity_total IN (0, 1)
  ),
  include_cost_total INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(include_cost_total) = 'integer' AND include_cost_total IN (0, 1)
  ),
  include_sales_total INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(include_sales_total) = 'integer' AND include_sales_total IN (0, 1)
  )
);

CREATE TABLE destination_aliases(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  raw_name TEXT NOT NULL,
  source_type TEXT NOT NULL,
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  UNIQUE(raw_name, source_type)
);

CREATE TABLE vehicle_rates(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
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
  active INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(active) = 'integer' AND active IN (0, 1)
  ),
  UNIQUE(destination_id, vehicle_type, effective_from),
  CHECK(effective_to IS NULL OR effective_to >= effective_from)
);

CREATE TRIGGER vehicle_rates_reject_overlap_insert
BEFORE INSERT ON vehicle_rates
WHEN EXISTS (
  SELECT 1
  FROM vehicle_rates AS existing
  WHERE existing.destination_id = NEW.destination_id
    AND existing.vehicle_type = NEW.vehicle_type
    AND NEW.effective_from <= COALESCE(existing.effective_to, '9999-12')
    AND COALESCE(NEW.effective_to, '9999-12') >= existing.effective_from
)
BEGIN
  SELECT RAISE(ABORT, 'vehicle rate effective period overlap');
END;

CREATE TRIGGER vehicle_rates_reject_overlap_update
BEFORE UPDATE ON vehicle_rates
WHEN EXISTS (
  SELECT 1
  FROM vehicle_rates AS existing
  WHERE existing.id <> OLD.id
    AND existing.destination_id = NEW.destination_id
    AND existing.vehicle_type = NEW.vehicle_type
    AND NEW.effective_from <= COALESCE(existing.effective_to, '9999-12')
    AND COALESCE(NEW.effective_to, '9999-12') >= existing.effective_from
)
BEGIN
  SELECT RAISE(ABORT, 'vehicle rate effective period overlap');
END;

CREATE TABLE report_groups(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  name TEXT NOT NULL UNIQUE,
  display_order INTEGER NOT NULL CHECK(
    typeof(display_order) = 'integer' AND display_order >= 0
  ),
  active INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(active) = 'integer' AND active IN (0, 1)
  )
);

CREATE TABLE report_group_members(
  group_id INTEGER NOT NULL CHECK(
    typeof(group_id) = 'integer' AND group_id > 0
  ) REFERENCES report_groups(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  display_order INTEGER NOT NULL CHECK(
    typeof(display_order) = 'integer' AND display_order >= 0
  ),
  include_quantity INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(include_quantity) = 'integer' AND include_quantity IN (0, 1)
  ),
  include_cost INTEGER NOT NULL DEFAULT 1 CHECK(
    typeof(include_cost) = 'integer' AND include_cost IN (0, 1)
  ),
  PRIMARY KEY(group_id, destination_id),
  UNIQUE(group_id, display_order)
);

CREATE TABLE monthly_plans(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  quantity_ea_text TEXT NOT NULL CHECK(
    typeof(quantity_ea_text) = 'text'
    AND quantity_ea_text = trim(quantity_ea_text)
    AND quantity_ea_text NOT GLOB '*[^0-9.]*'
    AND length(quantity_ea_text) - length(replace(quantity_ea_text, '.', '')) <= 1
    AND (
      (
        instr(quantity_ea_text, '.') = 0
        AND (
          quantity_ea_text = '0'
          OR substr(quantity_ea_text, 1, 1) BETWEEN '1' AND '9'
        )
      )
      OR (
        instr(quantity_ea_text, '.') > 1
        AND (
          substr(quantity_ea_text, 1, instr(quantity_ea_text, '.') - 1) = '0'
          OR substr(quantity_ea_text, 1, 1) BETWEEN '1' AND '9'
        )
        AND length(substr(quantity_ea_text, instr(quantity_ea_text, '.') + 1)) > 0
        AND substr(quantity_ea_text, -1, 1) BETWEEN '1' AND '9'
      )
    )
  ),
  cost_won INTEGER NOT NULL CHECK(typeof(cost_won) = 'integer' AND cost_won >= 0),
  representative_item TEXT,
  source_note TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, destination_id)
);

CREATE TABLE monthly_actual_quantities(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  quantity_ea_text TEXT NOT NULL CHECK(
    typeof(quantity_ea_text) = 'text'
    AND quantity_ea_text = trim(quantity_ea_text)
    AND quantity_ea_text NOT GLOB '*[^0-9.]*'
    AND length(quantity_ea_text) - length(replace(quantity_ea_text, '.', '')) <= 1
    AND (
      (
        instr(quantity_ea_text, '.') = 0
        AND (
          quantity_ea_text = '0'
          OR substr(quantity_ea_text, 1, 1) BETWEEN '1' AND '9'
        )
      )
      OR (
        instr(quantity_ea_text, '.') > 1
        AND (
          substr(quantity_ea_text, 1, instr(quantity_ea_text, '.') - 1) = '0'
          OR substr(quantity_ea_text, 1, 1) BETWEEN '1' AND '9'
        )
        AND length(substr(quantity_ea_text, instr(quantity_ea_text, '.') + 1)) > 0
        AND substr(quantity_ea_text, -1, 1) BETWEEN '1' AND '9'
      )
    )
  ),
  source_note TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, destination_id)
);

-- Report-cost precedence: aggregate committed transport_entries when they exist
-- for a destination/month; otherwise use this authoritative monthly total.
CREATE TABLE monthly_actual_costs(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER NOT NULL CHECK(
    typeof(destination_id) = 'integer' AND destination_id > 0
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  cost_won INTEGER NOT NULL CHECK(
    typeof(cost_won) = 'integer' AND cost_won >= 0
  ),
  source_type TEXT NOT NULL,
  source_note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, destination_id)
);

CREATE TABLE monthly_sales(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
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
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  source_type TEXT NOT NULL,
  source_filename TEXT NOT NULL,
  file_sha256 TEXT NOT NULL CHECK(
    typeof(file_sha256) = 'text'
    AND length(file_sha256) = 64
    AND file_sha256 NOT GLOB '*[^0-9a-fA-F]*'
  ),
  status TEXT NOT NULL,
  error_summary TEXT,
  imported_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(report_month, source_type, file_sha256),
  UNIQUE(id, report_month)
);

CREATE TABLE transport_entries(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  import_batch_id INTEGER NOT NULL CHECK(
    typeof(import_batch_id) = 'integer' AND import_batch_id > 0
  ),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  destination_id INTEGER CHECK(
    destination_id IS NULL OR (
      typeof(destination_id) = 'integer' AND destination_id > 0
    )
  ) REFERENCES destinations(id)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  unresolved_alias TEXT,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL CHECK(
    typeof(source_row) = 'integer' AND source_row > 0
  ),
  transport_day INTEGER NOT NULL CHECK(
    typeof(transport_day) = 'integer' AND transport_day BETWEEN 1 AND 31
  ),
  transport_type TEXT NOT NULL,
  vehicle_type TEXT,
  vehicle_driver_group TEXT,
  trip_count_text TEXT NOT NULL CHECK(
    typeof(trip_count_text) = 'text'
    AND trip_count_text = trim(trip_count_text)
    AND trip_count_text NOT GLOB '*[^0-9.]*'
    AND length(trip_count_text) - length(replace(trip_count_text, '.', '')) <= 1
    AND (
      (
        instr(trip_count_text, '.') = 0
        AND (
          trip_count_text = '0'
          OR substr(trip_count_text, 1, 1) BETWEEN '1' AND '9'
        )
      )
      OR (
        instr(trip_count_text, '.') > 1
        AND (
          substr(trip_count_text, 1, instr(trip_count_text, '.') - 1) = '0'
          OR substr(trip_count_text, 1, 1) BETWEEN '1' AND '9'
        )
        AND length(substr(trip_count_text, instr(trip_count_text, '.') + 1)) > 0
        AND substr(trip_count_text, -1, 1) BETWEEN '1' AND '9'
      )
    )
  ),
  unit_rate_won INTEGER CHECK(
    unit_rate_won IS NULL OR (
      typeof(unit_rate_won) = 'integer' AND unit_rate_won >= 0
    )
  ),
  cost_won INTEGER NOT NULL CHECK(typeof(cost_won) = 'integer' AND cost_won >= 0),
  source_note TEXT,
  FOREIGN KEY(import_batch_id, report_month)
    REFERENCES import_batches(id, report_month)
    ON UPDATE RESTRICT ON DELETE RESTRICT,
  UNIQUE(import_batch_id, source_sheet, source_row, transport_day, transport_type),
  CHECK(
    destination_id IS NOT NULL OR (
      unresolved_alias IS NOT NULL AND length(trim(unresolved_alias)) > 0
    )
  )
);

CREATE TABLE report_runs(
  id INTEGER PRIMARY KEY CHECK(typeof(id) = 'integer' AND id > 0),
  report_month TEXT NOT NULL CHECK(
    length(report_month) = 7
    AND substr(report_month, 1, 4) GLOB '[0-9][0-9][0-9][0-9]'
    AND substr(report_month, 5, 1) = '-'
    AND substr(report_month, 6, 2) BETWEEN '01' AND '12'
  ),
  status TEXT NOT NULL,
  is_locked INTEGER NOT NULL DEFAULT 0 CHECK(
    typeof(is_locked) = 'integer' AND is_locked IN (0, 1)
  ),
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
