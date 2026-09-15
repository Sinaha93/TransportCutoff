from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Destination:
    id: int
    name: str
    display_order: int
    active: bool
    required_for_report: bool
    representative_item: str | None
    include_quantity_total: bool
    include_cost_total: bool
    include_sales_total: bool


@dataclass(frozen=True, slots=True)
class DestinationAlias:
    id: int
    destination_id: int
    raw_name: str
    source_type: str


@dataclass(frozen=True, slots=True)
class VehicleRate:
    id: int
    destination_id: int
    vehicle_type: str
    rate_won: int
    effective_from_month: str
    effective_to_month: str | None
    source_note: str | None
    active: bool


@dataclass(frozen=True, slots=True)
class ReportGroup:
    id: int
    name: str
    display_order: int
    active: bool


@dataclass(frozen=True, slots=True)
class GroupMember:
    group_id: int
    destination_id: int
    name: str
    display_order: int
    include_quantity: bool
    include_cost: bool
