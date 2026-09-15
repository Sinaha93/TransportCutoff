import sqlite3
from dataclasses import FrozenInstanceError

import pytest

from app.db import Database
from app.repositories.masters import (
    DEFAULT_GROUPS,
    DEFAULT_TOTAL_RULES,
    DestinationInUseError,
    DuplicateAliasError,
    MasterDataError,
    MasterDataNotFoundError,
    MasterRepository,
    ValidationError,
    VehicleRateOverlapError,
    seed_default_masters,
)


@pytest.fixture
def repo(tmp_path):
    database = Database(tmp_path / "app.db")
    database.migrate()
    return MasterRepository(database)


def test_alias_resolves_to_standard_destination(repo):
    destination = repo.create_destination("포레시아 영천", display_order=10)
    repo.add_alias(destination.id, "포레시아-영천", "transport_workbook")
    assert repo.resolve_alias("포레시아-영천", "transport_workbook") == destination


def test_group_members_can_be_changed_without_formula_edits(repo):
    group = repo.create_group("영남권", display_order=20)
    members = [
        repo.create_destination(name, i)
        for i, name in enumerate(
            ["현대 울산", "포레시아 영천", "세종공업"], start=1
        )
    ]
    repo.replace_group_members(group.id, [m.id for m in members])
    assert [m.name for m in repo.group_members(group.id)] == [
        "현대 울산",
        "포레시아 영천",
        "세종공업",
    ]


def test_destination_names_are_trimmed_but_internal_punctuation_is_preserved(repo):
    destination = repo.create_destination("  포레시아-영천  ", 1)

    assert destination.name == "포레시아-영천"
    assert repo.get_destination(destination.id) == destination
    with pytest.raises(ValidationError, match="name"):
        repo.create_destination("   ", 2)


def test_destinations_have_deterministic_order_and_can_be_updated(repo):
    first = repo.create_destination("C", 1)
    second = repo.create_destination("A", 1)
    third = repo.create_destination("B", 2)

    assert [item.id for item in repo.list_destinations()] == [
        first.id,
        second.id,
        third.id,
    ]

    updated = repo.update_destination(
        third.id,
        name=" B-2 ",
        display_order=0,
        active=False,
        required_for_report=False,
        representative_item="ITEM-1",
        include_quantity_total=False,
        include_cost_total=True,
        include_sales_total=False,
    )
    assert updated.name == "B-2"
    assert updated.display_order == 0
    assert updated.active is False
    assert updated.required_for_report is False
    assert updated.representative_item == "ITEM-1"
    assert updated.include_quantity_total is False
    assert updated.include_cost_total is True
    assert updated.include_sales_total is False
    with pytest.raises(FrozenInstanceError):
        updated.name = "changed"


def test_alias_matching_uses_nfkc_and_trim_without_removing_punctuation(repo):
    destination = repo.create_destination("포레시아 영천", 1)
    alias = repo.add_alias(
        destination.id,
        "  포레시아－영천  ",
        "  ｔｒａｎｓｐｏｒｔ  ",
    )

    assert alias.raw_name == "포레시아-영천"
    assert alias.source_type == "transport"
    assert repo.resolve_alias(" 포레시아－영천 ", "ｔｒａｎｓｐｏｒｔ") == destination
    assert repo.resolve_alias("포레시아 영천", "transport") is None

    with pytest.raises(DuplicateAliasError, match="already exists"):
        repo.add_alias(destination.id, "포레시아-영천", "transport")


def test_aliases_can_be_listed_updated_and_removed(repo):
    first = repo.create_destination("First", 1)
    second = repo.create_destination("Second", 2)
    alias = repo.add_alias(first.id, "raw", "source")

    updated = repo.update_alias(
        alias.id,
        destination_id=second.id,
        raw_name=" updated ",
        source_type=" source-2 ",
    )
    assert updated.destination_id == second.id
    assert updated.raw_name == "updated"
    assert updated.source_type == "source-2"
    assert repo.list_aliases(second.id) == [updated]

    repo.remove_alias(updated.id)
    assert repo.list_aliases(second.id) == []


def test_alias_can_be_read_by_id_and_missing_returns_none(repo):
    destination = repo.create_destination("Destination", 1)
    alias = repo.add_alias(destination.id, "raw", "source")

    assert repo.get_alias(alias.id) == alias
    assert repo.get_alias(999999) is None


def test_group_order_flags_and_group_read_update_listing(repo):
    group = repo.create_group("  Group  ", 20)
    first = repo.create_destination("First", 1)
    second = repo.create_destination("Second", 2)

    repo.replace_group_members(
        group.id,
        [(second.id, False, True), (first.id, True, False)],
    )
    members = repo.group_members(group.id)
    assert [(item.name, item.display_order) for item in members] == [
        ("Second", 1),
        ("First", 2),
    ]
    assert [(item.include_quantity, item.include_cost) for item in members] == [
        (False, True),
        (True, False),
    ]

    updated = repo.update_group(group.id, name="Updated", display_order=2, active=False)
    assert repo.get_group(group.id) == updated
    assert repo.list_groups() == [updated]


def test_group_member_replacement_rolls_back_on_invalid_or_duplicate_members(repo):
    group = repo.create_group("Group", 1)
    first = repo.create_destination("First", 1)
    second = repo.create_destination("Second", 2)
    repo.replace_group_members(group.id, [first.id])

    with pytest.raises(MasterDataError):
        repo.replace_group_members(group.id, [second.id, 999999])
    assert [item.destination_id for item in repo.group_members(group.id)] == [first.id]

    with pytest.raises(ValidationError, match="duplicate"):
        repo.replace_group_members(group.id, [second.id, second.id])
    assert [item.destination_id for item in repo.group_members(group.id)] == [first.id]


def test_seed_is_idempotent_and_applies_known_group_and_total_rule(repo):
    seed_default_masters(repo)
    group = next(item for item in repo.list_groups() if item.name == "영남권")
    first_member_ids = [item.destination_id for item in repo.group_members(group.id)]
    first_destination_ids = [item.id for item in repo.list_destinations()]

    seed_default_masters(repo)

    assert DEFAULT_GROUPS == {"영남권": ["현대 울산", "포레시아 영천", "세종공업"]}
    assert DEFAULT_TOTAL_RULES == {
        "당진": {"quantity": False, "cost": True, "sales": False}
    }
    assert [item.name for item in repo.group_members(group.id)] == DEFAULT_GROUPS["영남권"]
    assert [item.destination_id for item in repo.group_members(group.id)] == first_member_ids
    assert [item.id for item in repo.list_destinations()] == first_destination_ids
    dangjin = next(item for item in repo.list_destinations() if item.name == "당진")
    assert (
        dangjin.include_quantity_total,
        dangjin.include_cost_total,
        dangjin.include_sales_total,
    ) == (False, True, False)


def test_seed_preserves_user_customized_membership_after_group_exists(repo):
    seed_default_masters(repo)
    group = next(item for item in repo.list_groups() if item.name == "영남권")
    custom = repo.create_destination("Custom", 99)
    repo.replace_group_members(group.id, [(custom.id, False, True)])

    seed_default_masters(repo)

    members = repo.group_members(group.id)
    assert [(item.name, item.include_quantity, item.include_cost) for item in members] == [
        ("Custom", False, True)
    ]


def test_rate_lookup_honors_boundaries_and_rates_can_be_updated(repo):
    destination = repo.create_destination("Destination", 1)
    first = repo.create_vehicle_rate(
        destination.id, "  5톤  ", 1000, "2026-01", "2026-06"
    )
    second = repo.create_vehicle_rate(
        destination.id, "5톤", 2000, "2026-07", None
    )

    assert repo.resolve_vehicle_rate(destination.id, "5톤", "2025-12") is None
    assert repo.resolve_vehicle_rate(destination.id, " 5톤 ", "2026-01") == first
    assert repo.resolve_vehicle_rate(destination.id, "5톤", "2026-06") == first
    assert repo.resolve_vehicle_rate(destination.id, "5톤", "2026-07") == second
    assert repo.resolve_vehicle_rate(destination.id, "5톤", "2099-12") == second

    updated = repo.update_vehicle_rate(first.id, rate_won=1250)
    assert updated.rate_won == 1250
    assert repo.list_vehicle_rates(destination.id) == [updated, second]


def test_rate_validation_and_overlap_errors_are_clear(repo):
    destination = repo.create_destination("Destination", 1)
    repo.create_vehicle_rate(destination.id, "5톤", 1000, "2026-01", "2026-06")

    with pytest.raises(VehicleRateOverlapError, match="overlap"):
        repo.create_vehicle_rate(destination.id, "5톤", 2000, "2026-06", None)
    with pytest.raises(ValidationError, match="rate_won"):
        repo.create_vehicle_rate(destination.id, "5톤", -1, "2027-01", None)
    with pytest.raises(ValidationError, match="YYYY-MM"):
        repo.create_vehicle_rate(destination.id, "5톤", 1, "2026-13", None)
    with pytest.raises(ValidationError, match="range"):
        repo.create_vehicle_rate(destination.id, "5톤", 1, "2026-08", "2026-07")


def test_vehicle_rate_can_be_read_and_deleted_without_affecting_other_rates(repo):
    destination = repo.create_destination("Destination", 1)
    removed = repo.create_vehicle_rate(
        destination.id, "5톤", 1000, "2026-01", "2026-06"
    )
    retained = repo.create_vehicle_rate(
        destination.id, "5톤", 2000, "2026-07", None
    )

    assert repo.get_vehicle_rate(removed.id) == removed
    assert repo.get_vehicle_rate(999999) is None

    repo.delete_vehicle_rate(removed.id)

    assert repo.get_vehicle_rate(removed.id) is None
    assert repo.list_vehicle_rates(destination.id) == [retained]
    with pytest.raises(MasterDataNotFoundError, match="Vehicle rate"):
        repo.delete_vehicle_rate(removed.id)


def test_group_delete_removes_members_but_preserves_destinations(repo):
    group = repo.create_group("Group", 1)
    first = repo.create_destination("First", 1)
    second = repo.create_destination("Second", 2)
    repo.replace_group_members(group.id, [first.id, second.id])

    repo.delete_group(group.id)

    assert repo.get_group(group.id) is None
    assert repo.group_members(group.id) == []
    assert repo.get_destination(first.id) == first
    assert repo.get_destination(second.id) == second
    with pytest.raises(MasterDataNotFoundError, match="Group"):
        repo.delete_group(group.id)


@pytest.mark.parametrize(
    "reference_type",
    [
        "alias",
        "rate",
        "group",
        "monthly_plan",
        "monthly_quantity",
        "monthly_cost",
        "transport",
    ],
)
def test_destination_deletion_is_protected_for_every_reference_class(
    repo, reference_type
):
    destination = repo.create_destination("Destination", 1)
    _add_destination_reference(repo.database, destination.id, reference_type)

    with pytest.raises(DestinationInUseError, match="referenced"):
        repo.delete_destination(destination.id)
    assert repo.get_destination(destination.id) == destination


def test_unreferenced_destination_can_be_deleted(repo):
    destination = repo.create_destination("Destination", 1)

    repo.delete_destination(destination.id)

    assert repo.get_destination(destination.id) is None


def test_repository_closes_connections_after_success(tmp_path):
    class TrackingDatabase(Database):
        last_connection = None

        def connect(self):
            connection = super().connect()
            self.last_connection = connection
            return connection

    database = TrackingDatabase(tmp_path / "app.db")
    database.migrate()
    repo = MasterRepository(database)

    repo.list_destinations()

    assert database.last_connection is not None
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        database.last_connection.execute("SELECT 1")


def _add_destination_reference(database, destination_id, reference_type):
    with database.connection() as connection:
        if reference_type == "alias":
            connection.execute(
                "INSERT INTO destination_aliases(raw_name, source_type, destination_id) "
                "VALUES (?, ?, ?)",
                ("raw", "source", destination_id),
            )
        elif reference_type == "rate":
            connection.execute(
                "INSERT INTO vehicle_rates"
                "(destination_id, vehicle_type, unit_rate_won, effective_from) "
                "VALUES (?, ?, ?, ?)",
                (destination_id, "5톤", 1000, "2026-01"),
            )
        elif reference_type == "group":
            group_id = connection.execute(
                "INSERT INTO report_groups(name, display_order) VALUES (?, ?)",
                ("Group", 1),
            ).lastrowid
            connection.execute(
                "INSERT INTO report_group_members"
                "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
                (group_id, destination_id, 1),
            )
        elif reference_type == "monthly_plan":
            connection.execute(
                "INSERT INTO monthly_plans"
                "(report_month, destination_id, quantity_ea_text, cost_won) "
                "VALUES (?, ?, ?, ?)",
                ("2026-01", destination_id, "1", 1000),
            )
        elif reference_type == "monthly_quantity":
            connection.execute(
                "INSERT INTO monthly_actual_quantities"
                "(report_month, destination_id, quantity_ea_text) VALUES (?, ?, ?)",
                ("2026-01", destination_id, "1"),
            )
        elif reference_type == "monthly_cost":
            connection.execute(
                "INSERT INTO monthly_actual_costs"
                "(report_month, destination_id, cost_won, source_type) "
                "VALUES (?, ?, ?, ?)",
                ("2026-01", destination_id, 1000, "test"),
            )
        elif reference_type == "transport":
            batch_id = connection.execute(
                "INSERT INTO import_batches"
                "(report_month, source_type, source_filename, file_sha256, status) "
                "VALUES (?, ?, ?, ?, ?)",
                ("2026-01", "test", "source.xlsx", "a" * 64, "committed"),
            ).lastrowid
            connection.execute(
                "INSERT INTO transport_entries"
                "(import_batch_id, report_month, destination_id, source_sheet, "
                "source_row, transport_day, transport_type, trip_count_text, cost_won) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    batch_id,
                    "2026-01",
                    destination_id,
                    "Sheet1",
                    2,
                    1,
                    "regular",
                    "1",
                    1000,
                ),
            )
        connection.commit()
