"""Riesgos (Gestión de Riesgos) integration port (CAP-053, H03)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol
from uuid import UUID


@dataclass(frozen=True)
class RiesgosLookup:
    """Risks link request keyed by stable identifier (CAP-053)."""

    expediente_id: UUID


@dataclass(frozen=True)
class RiesgosProject:
    """One project from the Riesgos system, returned by lookup."""

    riesgos_id: UUID
    expediente_id: UUID
    status: str
    payload: dict[str, Any]


class RiesgosPort(Protocol):
    """Gestión de Riesgos (CAP-053 camino feliz)."""

    async def lookup(self, request: RiesgosLookup, credential: str) -> RiesgosProject: ...


__all__ = ["RiesgosLookup", "RiesgosPort", "RiesgosProject"]
