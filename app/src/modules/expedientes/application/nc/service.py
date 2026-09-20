"""Service implementation for EXP-CAP-054 NC adapter (H04)."""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.nc._evidence import nc_event
from app.src.modules.expedientes.application.nc.command import (
    NCAuthorizationError,
    NCDependencyError,
    NCLookupCommand,
)
from app.src.modules.expedientes.ports.nc import NCLookup, NCRecord


class NCService:
    """Application-level NC adapter (CAP-054). Deny-by-default."""

    def __init__(self, *, nc: Any, audit_log: Any) -> None:
        self._nc = nc
        self._audit_log = audit_log

    async def lookup(self, command: NCLookupCommand) -> NCRecord:
        actor_id = _validated_actor(command.actor_id)
        if command.expediente_id is None and command.s4h_code is None:
            raise NCAuthorizationError("expediente_id or s4h_code is required (deny-by-default)")
        request = NCLookup(expediente_id=command.expediente_id, s4h_code=command.s4h_code)
        try:
            record = await self._nc.lookup(request, credential="")
        except Exception as exc:  # noqa: BLE001
            raise NCDependencyError("nc.lookup failed; caller UoW must roll back") from exc
        await self._audit_log.append(
            nc_event(
                actor_id=actor_id,
                target_id=record.expediente_id,
                payload={
                    "nc_id": str(record.nc_id),
                    "estado": record.estado,
                    "s4h_code": command.s4h_code or "",
                },
            )
        )
        return record


def _validated_actor(actor_id: UUID | None) -> UUID:
    if actor_id is None or actor_id == UUID(int=0):
        raise NCAuthorizationError("actor_id is required (deny-by-default)")
    return actor_id  # type: ignore[return-value]


__all__ = ["NCService"]
