from __future__ import annotations

import re
import sqlite3
import unicodedata
from collections.abc import Sequence
from typing import Final

from app.db import Database
from app.domain.models import (
    Destination,
    DestinationAlias,
    GroupMember,
    ReportGroup,
    VehicleRate,
)


DEFAULT_GROUPS: Final = {
    "영남권": ["현대 울산", "포레시아 영천", "세종공업"]
}
DEFAULT_TOTAL_RULES: Final = {
    "당진": {"quantity": False, "cost": True, "sales": False}
}

_MONTH = re.compile(r"^[0-9]{4}-(0[1-9]|1[0-2])$")


class _UnsetType:
    __slots__ = ()


_UNSET = _UnsetType()


class MasterDataError(Exception):
    """Base error for invalid or unsuccessful master-data operations."""


class MasterDataNotFoundError(MasterDataError):
    """Raised when a requested master-data record does not exist."""


class ValidationError(MasterDataError):
    """Raised when a master-data value is invalid."""


class DuplicateAliasError(MasterDataError):
    """Raised when a normalized alias and source pair already exists."""


class VehicleRateOverlapError(MasterDataError):
    """Raised when effective periods overlap for a destination and vehicle."""


class DestinationInUseError(MasterDataError):
    """Raised when deletion would remove a referenced destination."""


class MasterRepository:
    def __init__(self, database: Database) -> None:
        self.database = database

    def create_destination(
        self,
        name: str,
        display_order: int,
        *,
        active: bool = True,
        required_for_report: bool = True,
        representative_item: str | None = None,
        include_quantity_total: bool = True,
        include_cost_total: bool = True,
        include_sales_total: bool = True,
        initial_alias: tuple[str, str] | None = None,
    ) -> Destination:
        clean_name = _clean_text(name, "destination name")
        _validate_display_order(display_order)
        _validate_optional_text(representative_item, "representative_item")
        for field, value in (
            ("active", active),
            ("required_for_report", required_for_report),
            ("include_quantity_total", include_quantity_total),
            ("include_cost_total", include_cost_total),
            ("include_sales_total", include_sales_total),
        ):
            _validate_bool(value, field)
        if initial_alias is not None:
            initial_alias = tuple(_normalize_alias_text(value, "alias") for value in initial_alias)
            if len(initial_alias) != 2:
                raise ValidationError("Alias requires name and source type")
        parameters = (
            clean_name,
            display_order,
            representative_item,
            int(active),
            int(required_for_report),
            int(include_quantity_total),
            int(include_cost_total),
            int(include_sales_total),
        )
        with self.database.connection() as connection:
            try:
                row = connection.execute(
                    "INSERT INTO destinations"
                    "(name, display_order, representative_item, active, "
                    "required_for_report, include_quantity_total, "
                    "include_cost_total, include_sales_total) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING *",
                    parameters,
                ).fetchone()
                if initial_alias is not None:
                    connection.execute(
                        "INSERT INTO destination_aliases(destination_id, raw_name, source_type) VALUES (?, ?, ?)",
                        (row["id"], *initial_alias),
                    )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(
                    f"Destination could not be created: {error}"
                ) from error
        return _destination_from_row(row)

    def get_destination(self, destination_id: int) -> Destination | None:
        _validate_id(destination_id, "destination_id")
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM destinations WHERE id = ?", (destination_id,)
            ).fetchone()
        return None if row is None else _destination_from_row(row)

    def list_destinations(self) -> list[Destination]:
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT * FROM destinations ORDER BY display_order, id, name"
            ).fetchall()
        return [_destination_from_row(row) for row in rows]

    def update_destination(
        self,
        destination_id: int,
        *,
        name: str | _UnsetType = _UNSET,
        display_order: int | _UnsetType = _UNSET,
        active: bool | _UnsetType = _UNSET,
        required_for_report: bool | _UnsetType = _UNSET,
        representative_item: str | None | _UnsetType = _UNSET,
        include_quantity_total: bool | _UnsetType = _UNSET,
        include_cost_total: bool | _UnsetType = _UNSET,
        include_sales_total: bool | _UnsetType = _UNSET,
    ) -> Destination:
        _validate_id(destination_id, "destination_id")
        values: dict[str, object] = {}
        if name is not _UNSET:
            values["name"] = _clean_text(name, "destination name")
        if display_order is not _UNSET:
            _validate_display_order(display_order)
            values["display_order"] = display_order
        for column, value in (
            ("active", active),
            ("required_for_report", required_for_report),
            ("include_quantity_total", include_quantity_total),
            ("include_cost_total", include_cost_total),
            ("include_sales_total", include_sales_total),
        ):
            if value is not _UNSET:
                _validate_bool(value, column)
                values[column] = int(value)
        if representative_item is not _UNSET:
            _validate_optional_text(representative_item, "representative_item")
            values["representative_item"] = representative_item
        self._update_by_id("destinations", destination_id, values, "destination")
        result = self.get_destination(destination_id)
        if result is None:
            raise MasterDataError(f"Destination {destination_id} does not exist")
        return result

    def delete_destination(self, destination_id: int) -> None:
        _validate_id(destination_id, "destination_id")
        reference_tables = (
            "destination_aliases",
            "vehicle_rates",
            "report_group_members",
            "monthly_plans",
            "monthly_actual_quantities",
            "monthly_actual_costs",
            "transport_entries",
        )
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                referenced_by = [
                    table
                    for table in reference_tables
                    if connection.execute(
                        f"SELECT 1 FROM {table} WHERE destination_id = ? LIMIT 1",
                        (destination_id,),
                    ).fetchone()
                    is not None
                ]
                if referenced_by:
                    raise DestinationInUseError(
                        "Destination cannot be deleted because it is referenced by: "
                        + ", ".join(referenced_by)
                    )
                cursor = connection.execute(
                    "DELETE FROM destinations WHERE id = ?", (destination_id,)
                )
                if cursor.rowcount == 0:
                    raise MasterDataNotFoundError(
                        f"Destination {destination_id} does not exist"
                    )
                connection.commit()
            except (DestinationInUseError, MasterDataNotFoundError):
                connection.rollback()
                raise
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise DestinationInUseError(
                    "Destination cannot be deleted because it is referenced"
                ) from error

    def add_alias(
        self, destination_id: int, raw_name: str, source_type: str
    ) -> DestinationAlias:
        _validate_id(destination_id, "destination_id")
        clean_raw_name = _normalize_alias_text(raw_name, "raw_name")
        clean_source_type = _normalize_alias_text(source_type, "source_type")
        with self.database.connection() as connection:
            try:
                row = connection.execute(
                    "INSERT INTO destination_aliases"
                    "(destination_id, raw_name, source_type) VALUES (?, ?, ?) "
                    "RETURNING *",
                    (destination_id, clean_raw_name, clean_source_type),
                ).fetchone()
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                if "UNIQUE constraint failed" in str(error):
                    raise DuplicateAliasError(
                        "Alias already exists for this raw_name and source_type"
                    ) from error
                raise MasterDataError(f"Alias could not be created: {error}") from error
        return _alias_from_row(row)

    def update_alias(
        self,
        alias_id: int,
        *,
        destination_id: int | _UnsetType = _UNSET,
        raw_name: str | _UnsetType = _UNSET,
        source_type: str | _UnsetType = _UNSET,
    ) -> DestinationAlias:
        _validate_id(alias_id, "alias_id")
        values: dict[str, object] = {}
        if destination_id is not _UNSET:
            _validate_id(destination_id, "destination_id")
            values["destination_id"] = destination_id
        if raw_name is not _UNSET:
            values["raw_name"] = _normalize_alias_text(raw_name, "raw_name")
        if source_type is not _UNSET:
            values["source_type"] = _normalize_alias_text(source_type, "source_type")
        try:
            self._update_by_id("destination_aliases", alias_id, values, "alias")
        except MasterDataError as error:
            if "UNIQUE constraint failed" in str(error):
                raise DuplicateAliasError(
                    "Alias already exists for this raw_name and source_type"
                ) from error
            raise
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM destination_aliases WHERE id = ?", (alias_id,)
            ).fetchone()
        if row is None:
            raise MasterDataError(f"Alias {alias_id} does not exist")
        return _alias_from_row(row)

    def get_alias(self, alias_id: int) -> DestinationAlias | None:
        _validate_id(alias_id, "alias_id")
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM destination_aliases WHERE id = ?", (alias_id,)
            ).fetchone()
        return None if row is None else _alias_from_row(row)

    def remove_alias(self, alias_id: int) -> None:
        _validate_id(alias_id, "alias_id")
        with self.database.connection() as connection:
            cursor = connection.execute(
                "DELETE FROM destination_aliases WHERE id = ?", (alias_id,)
            )
            if cursor.rowcount == 0:
                connection.rollback()
                raise MasterDataNotFoundError(f"Alias {alias_id} does not exist")
            connection.commit()

    def list_aliases(
        self, destination_id: int | None = None
    ) -> list[DestinationAlias]:
        sql = "SELECT * FROM destination_aliases"
        parameters: tuple[object, ...] = ()
        if destination_id is not None:
            _validate_id(destination_id, "destination_id")
            sql += " WHERE destination_id = ?"
            parameters = (destination_id,)
        sql += " ORDER BY source_type, raw_name, id"
        with self.database.connection() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [_alias_from_row(row) for row in rows]

    def resolve_alias(self, raw_name: str, source_type: str) -> Destination | None:
        clean_raw_name = _normalize_alias_text(raw_name, "raw_name")
        clean_source_type = _normalize_alias_text(source_type, "source_type")
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT d.* FROM destination_aliases AS a "
                "JOIN destinations AS d ON d.id = a.destination_id "
                "WHERE a.raw_name = ? AND a.source_type = ?",
                (clean_raw_name, clean_source_type),
            ).fetchone()
        return None if row is None else _destination_from_row(row)

    def create_vehicle_rate(
        self,
        destination_id: int,
        vehicle_type: str,
        rate_won: int,
        effective_from_month: str,
        effective_to_month: str | None = None,
        *,
        source_note: str | None = None,
        active: bool = True,
    ) -> VehicleRate:
        _validate_id(destination_id, "destination_id")
        clean_vehicle_type = _clean_text(vehicle_type, "vehicle_type")
        _validate_rate(rate_won)
        _validate_month_range(effective_from_month, effective_to_month)
        _validate_bool(active, "active")
        _validate_optional_text(source_note, "source_note")
        with self.database.connection() as connection:
            try:
                row = connection.execute(
                    "INSERT INTO vehicle_rates"
                    "(destination_id, vehicle_type, unit_rate_won, effective_from, "
                    "effective_to, source_note, active) VALUES (?, ?, ?, ?, ?, ?, ?) "
                    "RETURNING *",
                    (
                        destination_id,
                        clean_vehicle_type,
                        rate_won,
                        effective_from_month,
                        effective_to_month,
                        source_note,
                        int(active),
                    ),
                ).fetchone()
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                _raise_rate_error(error)
        return _rate_from_row(row)

    def list_vehicle_rates(
        self, destination_id: int | None = None
    ) -> list[VehicleRate]:
        sql = "SELECT * FROM vehicle_rates"
        parameters: tuple[object, ...] = ()
        if destination_id is not None:
            _validate_id(destination_id, "destination_id")
            sql += " WHERE destination_id = ?"
            parameters = (destination_id,)
        sql += " ORDER BY destination_id, vehicle_type, effective_from, id"
        with self.database.connection() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [_rate_from_row(row) for row in rows]

    def get_vehicle_rate(self, rate_id: int) -> VehicleRate | None:
        _validate_id(rate_id, "rate_id")
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM vehicle_rates WHERE id = ?", (rate_id,)
            ).fetchone()
        return None if row is None else _rate_from_row(row)

    def delete_vehicle_rate(self, rate_id: int) -> None:
        _validate_id(rate_id, "rate_id")
        with self.database.connection() as connection:
            cursor = connection.execute(
                "DELETE FROM vehicle_rates WHERE id = ?", (rate_id,)
            )
            if cursor.rowcount == 0:
                connection.rollback()
                raise MasterDataNotFoundError(f"Vehicle rate {rate_id} does not exist")
            connection.commit()

    def update_vehicle_rate(
        self,
        rate_id: int,
        *,
        destination_id: int | _UnsetType = _UNSET,
        vehicle_type: str | _UnsetType = _UNSET,
        rate_won: int | _UnsetType = _UNSET,
        effective_from_month: str | _UnsetType = _UNSET,
        effective_to_month: str | None | _UnsetType = _UNSET,
        source_note: str | None | _UnsetType = _UNSET,
        active: bool | _UnsetType = _UNSET,
    ) -> VehicleRate:
        _validate_id(rate_id, "rate_id")
        current = self.get_vehicle_rate(rate_id)
        if current is None:
            raise MasterDataError(f"Vehicle rate {rate_id} does not exist")
        new_from = (
            current.effective_from_month
            if effective_from_month is _UNSET
            else effective_from_month
        )
        new_to = (
            current.effective_to_month
            if effective_to_month is _UNSET
            else effective_to_month
        )
        _validate_month_range(new_from, new_to)
        values: dict[str, object] = {}
        if destination_id is not _UNSET:
            _validate_id(destination_id, "destination_id")
            values["destination_id"] = destination_id
        if vehicle_type is not _UNSET:
            values["vehicle_type"] = _clean_text(vehicle_type, "vehicle_type")
        if rate_won is not _UNSET:
            _validate_rate(rate_won)
            values["unit_rate_won"] = rate_won
        if effective_from_month is not _UNSET:
            values["effective_from"] = effective_from_month
        if effective_to_month is not _UNSET:
            values["effective_to"] = effective_to_month
        if source_note is not _UNSET:
            _validate_optional_text(source_note, "source_note")
            values["source_note"] = source_note
        if active is not _UNSET:
            _validate_bool(active, "active")
            values["active"] = int(active)
        try:
            self._update_by_id("vehicle_rates", rate_id, values, "vehicle rate")
        except MasterDataError as error:
            if "overlap" in str(error):
                raise VehicleRateOverlapError(
                    "Vehicle rate effective period overlap"
                ) from error
            raise
        result = self.get_vehicle_rate(rate_id)
        if result is None:
            raise MasterDataError(f"Vehicle rate {rate_id} does not exist")
        return result

    def resolve_vehicle_rate(
        self, destination_id: int, vehicle_type: str, month: str
    ) -> VehicleRate | None:
        _validate_id(destination_id, "destination_id")
        clean_vehicle_type = _clean_text(vehicle_type, "vehicle_type")
        _validate_month(month)
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT * FROM vehicle_rates "
                "WHERE destination_id = ? AND vehicle_type = ? AND active = 1 "
                "AND effective_from <= ? "
                "AND (effective_to IS NULL OR effective_to >= ?) "
                "ORDER BY effective_from, id",
                (destination_id, clean_vehicle_type, month, month),
            ).fetchall()
        if len(rows) > 1:
            raise MasterDataError("Multiple applicable vehicle rates were found")
        return None if not rows else _rate_from_row(rows[0])

    def create_group(
        self, name: str, display_order: int, *, active: bool = True
    ) -> ReportGroup:
        clean_name = _clean_text(name, "group name")
        _validate_display_order(display_order)
        _validate_bool(active, "active")
        with self.database.connection() as connection:
            try:
                row = connection.execute(
                    "INSERT INTO report_groups(name, display_order, active) "
                    "VALUES (?, ?, ?) RETURNING *",
                    (clean_name, display_order, int(active)),
                ).fetchone()
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(f"Group could not be created: {error}") from error
        return _group_from_row(row)

    def get_group(self, group_id: int) -> ReportGroup | None:
        _validate_id(group_id, "group_id")
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM report_groups WHERE id = ?", (group_id,)
            ).fetchone()
        return None if row is None else _group_from_row(row)

    def save_group_form(
        self, group_id: int | None, name: str, display_order: int,
        active: bool, members: Sequence[tuple[int, bool, bool]],
    ) -> None:
        """Atomically save group metadata and ordered member inclusion rules."""
        clean_name = _clean_text(name, "group name")
        _validate_display_order(display_order)
        _validate_bool(active, "active")
        normalized = [_normalize_group_member(member) for member in members]
        ids = [member[0] for member in normalized]
        if len(ids) != len(set(ids)):
            raise ValidationError("그룹 납품처가 중복되었습니다.")
        if group_id is not None:
            _validate_id(group_id, "group_id")
        with self.database.connection() as connection, connection:
            connection.execute("BEGIN IMMEDIATE")
            if group_id is None:
                group_id = connection.execute(
                    "INSERT INTO report_groups(name, display_order, active) VALUES (?, ?, ?)",
                    (clean_name, display_order, int(active)),
                ).lastrowid
            elif connection.execute(
                "UPDATE report_groups SET name = ?, display_order = ?, active = ? WHERE id = ?",
                (clean_name, display_order, int(active), group_id),
            ).rowcount != 1:
                raise MasterDataNotFoundError("그룹을 찾을 수 없습니다.")
            connection.execute("DELETE FROM report_group_members WHERE group_id = ?", (group_id,))
            connection.executemany(
                "INSERT INTO report_group_members(group_id, destination_id, display_order, include_quantity, include_cost) VALUES (?, ?, ?, ?, ?)",
                [(group_id, destination_id, order, int(quantity), int(cost))
                 for order, (destination_id, quantity, cost) in enumerate(normalized, 1)],
            )

    def list_groups(self) -> list[ReportGroup]:
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT * FROM report_groups ORDER BY display_order, id, name"
            ).fetchall()
        return [_group_from_row(row) for row in rows]

    def update_group(
        self,
        group_id: int,
        *,
        name: str | _UnsetType = _UNSET,
        display_order: int | _UnsetType = _UNSET,
        active: bool | _UnsetType = _UNSET,
    ) -> ReportGroup:
        _validate_id(group_id, "group_id")
        values: dict[str, object] = {}
        if name is not _UNSET:
            values["name"] = _clean_text(name, "group name")
        if display_order is not _UNSET:
            _validate_display_order(display_order)
            values["display_order"] = display_order
        if active is not _UNSET:
            _validate_bool(active, "active")
            values["active"] = int(active)
        self._update_by_id("report_groups", group_id, values, "group")
        result = self.get_group(group_id)
        if result is None:
            raise MasterDataError(f"Group {group_id} does not exist")
        return result

    def delete_group(self, group_id: int) -> None:
        _validate_id(group_id, "group_id")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN")
                if connection.execute(
                    "SELECT 1 FROM report_groups WHERE id = ?", (group_id,)
                ).fetchone() is None:
                    raise MasterDataNotFoundError(
                        f"Group {group_id} does not exist"
                    )
                connection.execute(
                    "DELETE FROM report_group_members WHERE group_id = ?", (group_id,)
                )
                connection.execute(
                    "DELETE FROM report_groups WHERE id = ?", (group_id,)
                )
                connection.commit()
            except MasterDataNotFoundError:
                connection.rollback()
                raise
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(f"Group could not be deleted: {error}") from error

    def replace_group_members(
        self,
        group_id: int,
        members: Sequence[int | tuple[int, bool, bool]],
    ) -> None:
        _validate_id(group_id, "group_id")
        normalized = [_normalize_group_member(member) for member in members]
        destination_ids = [member[0] for member in normalized]
        if len(destination_ids) != len(set(destination_ids)):
            raise ValidationError("Group members contain a duplicate destination")
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN")
                if connection.execute(
                    "SELECT 1 FROM report_groups WHERE id = ?", (group_id,)
                ).fetchone() is None:
                    raise MasterDataError(f"Group {group_id} does not exist")
                connection.execute(
                    "DELETE FROM report_group_members WHERE group_id = ?", (group_id,)
                )
                connection.executemany(
                    "INSERT INTO report_group_members"
                    "(group_id, destination_id, display_order, "
                    "include_quantity, include_cost) VALUES (?, ?, ?, ?, ?)",
                    (
                        (
                            group_id,
                            destination_id,
                            display_order,
                            int(include_quantity),
                            int(include_cost),
                        )
                        for display_order, (
                            destination_id,
                            include_quantity,
                            include_cost,
                        ) in enumerate(normalized, start=1)
                    ),
                )
                connection.commit()
            except MasterDataError:
                connection.rollback()
                raise
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(
                    f"Group members could not be replaced: {error}"
                ) from error

    def group_members(self, group_id: int) -> list[GroupMember]:
        _validate_id(group_id, "group_id")
        with self.database.connection() as connection:
            rows = connection.execute(
                "SELECT gm.*, d.name FROM report_group_members AS gm "
                "JOIN destinations AS d ON d.id = gm.destination_id "
                "WHERE gm.group_id = ? ORDER BY gm.display_order, gm.destination_id",
                (group_id,),
            ).fetchall()
        return [
            GroupMember(
                group_id=int(row["group_id"]),
                destination_id=int(row["destination_id"]),
                name=str(row["name"]),
                display_order=int(row["display_order"]),
                include_quantity=bool(row["include_quantity"]),
                include_cost=bool(row["include_cost"]),
            )
            for row in rows
        ]

    def seed_defaults(self) -> None:
        with self.database.connection() as connection:
            try:
                connection.execute("BEGIN IMMEDIATE")
                destination_rows = connection.execute(
                    "SELECT id, name, display_order FROM destinations"
                ).fetchall()
                destination_ids = {
                    str(row["name"]): int(row["id"]) for row in destination_rows
                }
                created_destination_names: set[str] = set()
                next_destination_order = (
                    max(
                        (int(row["display_order"]) for row in destination_rows),
                        default=0,
                    )
                    + 1
                )
                required_names = [
                    name for members in DEFAULT_GROUPS.values() for name in members
                ] + list(DEFAULT_TOTAL_RULES)
                for name in dict.fromkeys(required_names):
                    if name in destination_ids:
                        continue
                    cursor = connection.execute(
                        "INSERT INTO destinations(name, display_order) VALUES (?, ?)",
                        (name, next_destination_order),
                    )
                    destination_ids[name] = int(cursor.lastrowid)
                    created_destination_names.add(name)
                    next_destination_order += 1

                for name, rules in DEFAULT_TOTAL_RULES.items():
                    if name not in created_destination_names:
                        continue
                    connection.execute(
                        "UPDATE destinations SET include_quantity_total = ?, "
                        "include_cost_total = ?, include_sales_total = ? WHERE id = ?",
                        (
                            int(rules["quantity"]),
                            int(rules["cost"]),
                            int(rules["sales"]),
                            destination_ids[name],
                        ),
                    )

                group_rows = connection.execute(
                    "SELECT id, name, display_order FROM report_groups"
                ).fetchall()
                existing_group_names = {str(row["name"]) for row in group_rows}
                next_group_order = (
                    max(
                        (int(row["display_order"]) for row in group_rows), default=0
                    )
                    + 1
                )
                for name, member_names in DEFAULT_GROUPS.items():
                    if name in existing_group_names:
                        continue
                    group_id = connection.execute(
                        "INSERT INTO report_groups(name, display_order) VALUES (?, ?)",
                        (name, next_group_order),
                    ).lastrowid
                    next_group_order += 1
                    connection.executemany(
                        "INSERT INTO report_group_members"
                        "(group_id, destination_id, display_order) VALUES (?, ?, ?)",
                        (
                            (group_id, destination_ids[member_name], display_order)
                            for display_order, member_name in enumerate(
                                member_names, start=1
                            )
                        ),
                    )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(f"Defaults could not be seeded: {error}") from error
            except Exception:
                connection.rollback()
                raise

    def _update_by_id(
        self, table: str, record_id: int, values: dict[str, object], label: str
    ) -> None:
        if not values:
            return
        assignments = ", ".join(f"{column} = ?" for column in values)
        parameters = (*values.values(), record_id)
        with self.database.connection() as connection:
            try:
                cursor = connection.execute(
                    f"UPDATE {table} SET {assignments} WHERE id = ?", parameters
                )
                if cursor.rowcount == 0:
                    raise MasterDataError(f"{label.title()} {record_id} does not exist")
                connection.commit()
            except MasterDataError:
                connection.rollback()
                raise
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(
                    f"{label.title()} could not be updated: {error}"
                ) from error


def seed_default_masters(repository: MasterRepository) -> None:
    repository.seed_defaults()


def _clean_text(value: object, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{field} must not be empty")
    return value.strip()


def _normalize_alias_text(value: object, field: str) -> str:
    if not isinstance(value, str):
        raise ValidationError(f"{field} must be text")
    return _clean_text(unicodedata.normalize("NFKC", value), field)


def _validate_display_order(value: object) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValidationError("display_order must be a nonnegative integer")


def _validate_id(value: object, field: str) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValidationError(f"{field} must be a positive integer")


def _validate_bool(value: object, field: str) -> None:
    if not isinstance(value, bool):
        raise ValidationError(f"{field} must be a boolean")


def _validate_optional_text(value: object, field: str) -> None:
    if value is not None and not isinstance(value, str):
        raise ValidationError(f"{field} must be text or None")


def _validate_rate(value: object) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValidationError("rate_won must be a nonnegative integer")


def _validate_month(value: object) -> None:
    if not isinstance(value, str) or _MONTH.fullmatch(value) is None:
        raise ValidationError("effective month must use YYYY-MM")


def _validate_month_range(start: object, end: object) -> None:
    _validate_month(start)
    if end is not None:
        _validate_month(end)
        if end < start:
            raise ValidationError("effective month range ends before it starts")


def _normalize_group_member(
    member: int | tuple[int, bool, bool],
) -> tuple[int, bool, bool]:
    if isinstance(member, bool):
        raise ValidationError("Group member destination id must be an integer")
    if isinstance(member, int):
        _validate_id(member, "destination_id")
        return member, True, True
    if not isinstance(member, tuple) or len(member) != 3:
        raise ValidationError(
            "Group member must be an id or (id, include_quantity, include_cost)"
        )
    destination_id, include_quantity, include_cost = member
    _validate_id(destination_id, "destination_id")
    if not isinstance(include_quantity, bool) or not isinstance(include_cost, bool):
        raise ValidationError("Group member inclusion flags must be booleans")
    return destination_id, include_quantity, include_cost


def _raise_rate_error(error: sqlite3.IntegrityError) -> None:
    if "overlap" in str(error):
        raise VehicleRateOverlapError(
            "Vehicle rate effective period overlap"
        ) from error
    raise MasterDataError(f"Vehicle rate could not be created: {error}") from error


def _destination_from_row(row: sqlite3.Row) -> Destination:
    return Destination(
        id=int(row["id"]),
        name=str(row["name"]),
        display_order=int(row["display_order"]),
        active=bool(row["active"]),
        required_for_report=bool(row["required_for_report"]),
        representative_item=row["representative_item"],
        include_quantity_total=bool(row["include_quantity_total"]),
        include_cost_total=bool(row["include_cost_total"]),
        include_sales_total=bool(row["include_sales_total"]),
    )


def _alias_from_row(row: sqlite3.Row) -> DestinationAlias:
    return DestinationAlias(
        id=int(row["id"]),
        destination_id=int(row["destination_id"]),
        raw_name=str(row["raw_name"]),
        source_type=str(row["source_type"]),
    )


def _rate_from_row(row: sqlite3.Row) -> VehicleRate:
    return VehicleRate(
        id=int(row["id"]),
        destination_id=int(row["destination_id"]),
        vehicle_type=str(row["vehicle_type"]),
        rate_won=int(row["unit_rate_won"]),
        effective_from_month=str(row["effective_from"]),
        effective_to_month=row["effective_to"],
        source_note=row["source_note"],
        active=bool(row["active"]),
    )


def _group_from_row(row: sqlite3.Row) -> ReportGroup:
    return ReportGroup(
        id=int(row["id"]),
        name=str(row["name"]),
        display_order=int(row["display_order"]),
        active=bool(row["active"]),
    )
