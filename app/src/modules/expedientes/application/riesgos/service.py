"""Service implementation for EXP-CAP-053 Riesgos adapter (H03).

``RiesgosService.lookup`` wraps ``RiesgosPort.lookup`` with deny-by-default
+ one audit event per decision. Duck-typed to keep cross-module imports
out of the application layer (``scripts/check_layers.py``).
"""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.riesgos._evidence import riesgos_event
from app.src.modules.expedientes.application.riesgos.command import (
    RiesgosAuthorizationError,
    RiesgosDependencyError,
    RiesgosLookupCommand,
)
from app.src.modules.expedientes.ports.riesgos import RiesgosLookup


class RiesgosProject:
    """Result of a Riesgos lookup (kept local; mirror RiesgosProject
    port dataclass without importing the port)."""

    def __init__(
        self, riesgos_id: UUID, expediente_id: UUID, status: str, payload: dict[str, Any]
    ) -> None:
        self.riesgos_id = riesgos_id
        self.expediente_id = expediente_id
        self.status = status
        self.payload = payload


class RiesgosService:
    """Application-level Riesgos adapter (CAP-053). Deny-by-default."""

    def __init__(self, *, riesgos: Any, audit_log: Any) -> None:
        self._riesgos = riesgos
        self._audit_log = audit_log

    async def lookup(self, command: RiesgosLookupCommand) -> RiesgosProject:
        actor_id = _validated_actor(command.actor_id)
        request = RiesgosLookup(expediente_id=command.expediente_id)
        try:
            record = await self._riesgos.lookup(request, credential="")
        except Exception as exc:  # noqa: BLE001
            raise RiesgosDependencyError(
                "riesgos.lookup failed; caller UoW must roll back"
            ) from exc
        await self._audit_log.append(
            riesgos_event(
                actor_id=actor_id,
                target_id=record.expediente_id,
                payload={"riesgos_id": str(record.riesgos_id), "status": record.status},
            )
        )
        return RiesgosProject(
            riesgos_id=record.riesgos_id,
            expediente_id=record.expediente_id,
            status=record.status,
            payload=record.payload,
        )


def _validated_actor(actor_id: UUID | None) -> UUID:
    if actor_id is None or actor_id == UUID(int=0):
        raise RiesgosAuthorizationError("actor_id is required (deny-by-default)")
    return actor_id


__all__ = ["RiesgosProject", "RiesgosService"]
