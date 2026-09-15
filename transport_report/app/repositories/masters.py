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

_MONTH = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
_UNSET = object()


class MasterDataError(Exception):
    """Base error for invalid or unsuccessful master-data operations."""


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
    ) -> Destination:
        clean_name = _clean_text(name, "destination name")
        _validate_display_order(display_order)
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
                cursor = connection.execute(
                    "INSERT INTO destinations"
                    "(name, display_order, representative_item, active, "
                    "required_for_report, include_quantity_total, "
                    "include_cost_total, include_sales_total) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    parameters,
                )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(
                    f"Destination could not be created: {error}"
                ) from error
            row = connection.execute(
                "SELECT * FROM destinations WHERE id = ?", (cursor.lastrowid,)
            ).fetchone()
        return _destination_from_row(row)

    def get_destination(self, destination_id: int) -> Destination | None:
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
        name: str | object = _UNSET,
        display_order: int | object = _UNSET,
        active: bool | object = _UNSET,
        required_for_report: bool | object = _UNSET,
        representative_item: str | None | object = _UNSET,
        include_quantity_total: bool | object = _UNSET,
        include_cost_total: bool | object = _UNSET,
        include_sales_total: bool | object = _UNSET,
    ) -> Destination:
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
                values[column] = int(bool(value))
        if representative_item is not _UNSET:
            values["representative_item"] = representative_item
        self._update_by_id("destinations", destination_id, values, "destination")
        result = self.get_destination(destination_id)
        if result is None:
            raise MasterDataError(f"Destination {destination_id} does not exist")
        return result

    def delete_destination(self, destination_id: int) -> None:
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
            connection.execute(
                "DELETE FROM destinations WHERE id = ?", (destination_id,)
            )
            connection.commit()

    def add_alias(
        self, destination_id: int, raw_name: str, source_type: str
    ) -> DestinationAlias:
        clean_raw_name = _normalize_alias_text(raw_name, "raw_name")
        clean_source_type = _normalize_alias_text(source_type, "source_type")
        with self.database.connection() as connection:
            try:
                cursor = connection.execute(
                    "INSERT INTO destination_aliases"
                    "(destination_id, raw_name, source_type) VALUES (?, ?, ?)",
                    (destination_id, clean_raw_name, clean_source_type),
                )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                if "UNIQUE constraint failed" in str(error):
                    raise DuplicateAliasError(
                        "Alias already exists for this raw_name and source_type"
                    ) from error
                raise MasterDataError(f"Alias could not be created: {error}") from error
            row = connection.execute(
                "SELECT * FROM destination_aliases WHERE id = ?", (cursor.lastrowid,)
            ).fetchone()
        return _alias_from_row(row)

    def update_alias(
        self,
        alias_id: int,
        *,
        destination_id: int | object = _UNSET,
        raw_name: str | object = _UNSET,
        source_type: str | object = _UNSET,
    ) -> DestinationAlias:
        values: dict[str, object] = {}
        if destination_id is not _UNSET:
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

    def remove_alias(self, alias_id: int) -> None:
        with self.database.connection() as connection:
            connection.execute(
                "DELETE FROM destination_aliases WHERE id = ?", (alias_id,)
            )
            connection.commit()

    def list_aliases(
        self, destination_id: int | None = None
    ) -> list[DestinationAlias]:
        sql = "SELECT * FROM destination_aliases"
        parameters: tuple[object, ...] = ()
        if destination_id is not None:
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
        clean_vehicle_type = _clean_text(vehicle_type, "vehicle_type")
        _validate_rate(rate_won)
        _validate_month_range(effective_from_month, effective_to_month)
        with self.database.connection() as connection:
            try:
                cursor = connection.execute(
                    "INSERT INTO vehicle_rates"
                    "(destination_id, vehicle_type, unit_rate_won, effective_from, "
                    "effective_to, source_note, active) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (
                        destination_id,
                        clean_vehicle_type,
                        rate_won,
                        effective_from_month,
                        effective_to_month,
                        source_note,
                        int(active),
                    ),
                )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                _raise_rate_error(error)
            row = connection.execute(
                "SELECT * FROM vehicle_rates WHERE id = ?", (cursor.lastrowid,)
            ).fetchone()
        return _rate_from_row(row)

    def list_vehicle_rates(
        self, destination_id: int | None = None
    ) -> list[VehicleRate]:
        sql = "SELECT * FROM vehicle_rates"
        parameters: tuple[object, ...] = ()
        if destination_id is not None:
            sql += " WHERE destination_id = ?"
            parameters = (destination_id,)
        sql += " ORDER BY destination_id, vehicle_type, effective_from, id"
        with self.database.connection() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [_rate_from_row(row) for row in rows]

    def update_vehicle_rate(
        self,
        rate_id: int,
        *,
        destination_id: int | object = _UNSET,
        vehicle_type: str | object = _UNSET,
        rate_won: int | object = _UNSET,
        effective_from_month: str | object = _UNSET,
        effective_to_month: str | None | object = _UNSET,
        source_note: str | None | object = _UNSET,
        active: bool | object = _UNSET,
    ) -> VehicleRate:
        current = self._get_vehicle_rate(rate_id)
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
            values["source_note"] = source_note
        if active is not _UNSET:
            values["active"] = int(bool(active))
        try:
            self._update_by_id("vehicle_rates", rate_id, values, "vehicle rate")
        except MasterDataError as error:
            if "overlap" in str(error):
                raise VehicleRateOverlapError(
                    "Vehicle rate effective period overlap"
                ) from error
            raise
        result = self._get_vehicle_rate(rate_id)
        if result is None:
            raise MasterDataError(f"Vehicle rate {rate_id} does not exist")
        return result

    def resolve_vehicle_rate(
        self, destination_id: int, vehicle_type: str, month: str
    ) -> VehicleRate | None:
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
        with self.database.connection() as connection:
            try:
                cursor = connection.execute(
                    "INSERT INTO report_groups(name, display_order, active) "
                    "VALUES (?, ?, ?)",
                    (clean_name, display_order, int(active)),
                )
                connection.commit()
            except sqlite3.IntegrityError as error:
                connection.rollback()
                raise MasterDataError(f"Group could not be created: {error}") from error
            row = connection.execute(
                "SELECT * FROM report_groups WHERE id = ?", (cursor.lastrowid,)
            ).fetchone()
        return _group_from_row(row)

    def get_group(self, group_id: int) -> ReportGroup | None:
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM report_groups WHERE id = ?", (group_id,)
            ).fetchone()
        return None if row is None else _group_from_row(row)

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
        name: str | object = _UNSET,
        display_order: int | object = _UNSET,
        active: bool | object = _UNSET,
    ) -> ReportGroup:
        values: dict[str, object] = {}
        if name is not _UNSET:
            values["name"] = _clean_text(name, "group name")
        if display_order is not _UNSET:
            _validate_display_order(display_order)
            values["display_order"] = display_order
        if active is not _UNSET:
            values["active"] = int(bool(active))
        self._update_by_id("report_groups", group_id, values, "group")
        result = self.get_group(group_id)
        if result is None:
            raise MasterDataError(f"Group {group_id} does not exist")
        return result

    def replace_group_members(
        self,
        group_id: int,
        members: Sequence[int | tuple[int, bool, bool]],
    ) -> None:
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

    def _get_vehicle_rate(self, rate_id: int) -> VehicleRate | None:
        with self.database.connection() as connection:
            row = connection.execute(
                "SELECT * FROM vehicle_rates WHERE id = ?", (rate_id,)
            ).fetchone()
        return None if row is None else _rate_from_row(row)

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
    existing_destinations = {
        destination.name: destination for destination in repository.list_destinations()
    }
    next_destination_order = (
        max((item.display_order for item in existing_destinations.values()), default=0) + 1
    )
    required_names = [
        name for members in DEFAULT_GROUPS.values() for name in members
    ] + list(DEFAULT_TOTAL_RULES)
    for name in dict.fromkeys(required_names):
        if name not in existing_destinations:
            existing_destinations[name] = repository.create_destination(
                name, next_destination_order
            )
            next_destination_order += 1

    for name, rules in DEFAULT_TOTAL_RULES.items():
        destination = existing_destinations[name]
        existing_destinations[name] = repository.update_destination(
            destination.id,
            include_quantity_total=rules["quantity"],
            include_cost_total=rules["cost"],
            include_sales_total=rules["sales"],
        )

    existing_groups = {group.name: group for group in repository.list_groups()}
    next_group_order = max(
        (item.display_order for item in existing_groups.values()), default=0
    ) + 1
    for name, member_names in DEFAULT_GROUPS.items():
        if name in existing_groups:
            continue
        group = repository.create_group(name, next_group_order)
        next_group_order += 1
        repository.replace_group_members(
            group.id,
            [existing_destinations[member_name].id for member_name in member_names],
        )


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
        return member, True, True
    if not isinstance(member, tuple) or len(member) != 3:
        raise ValidationError(
            "Group member must be an id or (id, include_quantity, include_cost)"
        )
    destination_id, include_quantity, include_cost = member
    if isinstance(destination_id, bool) or not isinstance(destination_id, int):
        raise ValidationError("Group member destination id must be an integer")
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
