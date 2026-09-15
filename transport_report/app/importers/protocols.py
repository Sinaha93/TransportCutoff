from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Protocol, runtime_checkable


@dataclass(frozen=True, slots=True)
class QuantityRecord:
    report_month: str
    destination_id: int
    quantity_ea: Decimal
    source: str


@runtime_checkable
class QuantityProvider(Protocol):
    def load(self, report_month: str) -> list[QuantityRecord]: ...
