"""No Conformidades (NC) integration port (CAP-054, H04)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol
from uuid import UUID


@dataclass(frozen=True)
class NCLookup:
    """NC query keyed by expediente or S4H code (CAP-054 camino feliz)."""

    expediente_id: UUID | None = None
    s4h_code: str | None = None


@dataclass(frozen=True)
class NCRecord:
    """NC state for one expediente (CAP-054 consulta)."""

    nc_id: UUID
    expediente_id: UUID
    estado: str  # "abierta" / "cerrada" / "en_progreso" etc.
    payload: dict[str, Any]


class NCPort(Protocol):
    """No Conformidades lookup."""

    async def lookup(self, request: NCLookup, credential: str) -> NCRecord: ...


__all__ = ["NCLookup", "NCPort", "NCRecord"]
