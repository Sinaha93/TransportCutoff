# Persisted import-batch selection implementation plan

**Goal:** Persist exactly one selected current transport import per report month and source, and make report validation consume that persisted provenance.

**Contract:** Migration 005 adds strict `is_current` storage and a partial unique index. Existing groups select their greatest batch id among statuses `imported` and `imported_with_errors`; groups with no eligible batch remain unselected. New successful imports replace selection atomically, duplicate hashes preserve their existing selection state, and readers expose current batches explicitly.

## Task 1: Migration behavior

- Add failing tests for fresh migration, populated upgrade, eligible-status backfill, row preservation, strict boolean/index enforcement, and transactional rollback.
- Add `005_import_batch_selection.sql` with the column, deterministic backfill, and partial unique index.
- Move the documented future ERP item-mapping migration from 005 to 006.

## Task 2: Repository lifecycle

- Add failing tests for first selection, corrected replacement, same-current retry, superseded-hash retry, locked-month preservation, rollback preservation, and current-only query.
- Extend import results with persisted selection state and add an immutable current-batch snapshot.
- Switch selection only after all batch and entry writes have succeeded, inside the existing transaction.

## Task 3: Validation provenance

- Add failing tests for wrong-month current batches and case-insensitive SHA identity.
- Canonicalize valid SHA-256 values to lowercase in the frozen input model.
- Validate every current batch month independently; limit duplicate grouping to current batches in the requested month.

## Task 4: Verification and review

- Run focused validation, database, and importer tests with warnings as errors.
- Run the full suite with warnings as errors, compile all application modules, inspect the diff, and request an independent review.
- Fix review findings, rerun verification, then commit exactly `fix: persist current import batches` with a clean worktree.
