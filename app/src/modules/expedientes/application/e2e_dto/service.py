"""E2E DTO exchange service for EXP-CAP-033."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.e2e_dto.command import (
    SUPPORTED_API_VERSION,
    E2EDtoAuthorizationError,
    E2EDtoCommand,
    E2EDtoError,
    E2EDtoResult,
    E2EDtoValidationError,
    ExpedienteDTO,
    from_aggregate,
)
from app.src.modules.expedientes.ports.audit_log import AuditLogPort


class E2EDtoService:
    """Stable DTO exchange for the E2E channel (CAP-033).

    The service is a thin layer over the domain aggregate: it serialises
    the aggregate to a versioned DTO, validates the inbound version, and
    emits one audit event per exchange. The actual JSON rendering is
    the adapter concern; the use case only produces the typed DTO.
    """

    def __init__(
        self,
        *,
        audit_log: AuditLogPort,
        permissions: set[str] | None = None,
        supported_version: str = SUPPORTED_API_VERSION,
    ) -> None:
        self._audit_log = audit_log
        self._permissions: set[str] = set(permissions or ())
        self._supported_version = supported_version

    def grant(self, permission: str) -> None:
        self._permissions.add(permission)

    async def exchange(self, command: E2EDtoCommand) -> E2EDtoResult:
        self._check_actor(command.actor_id)
        if command.api_version != self._supported_version:
            raise E2EDtoValidationError(
                f"apiVersion {command.api_version!r} not supported; "
                f"expected {self._supported_version!r}"
            )
        dto = from_aggregate(command.expediente, api_version=command.api_version)
        try:
            reconstructed = dto.to_aggregate()
        except E2EDtoValidationError:
            raise
        except E2EDtoError:
            # E2EDtoError is the domain failure path (e.g. invalid
            # estado/tipo combination). Treat it as a transient error
            # the caller can retry with a different aggregate.
            raise
        except Exception as exc:
            raise E2EDtoError(f"dto exchange failed: {exc}") from exc
        await self._audit_log.append(_exchange_event(command.actor_id, dto, datetime.now(UTC)))
        return E2EDtoResult(dto=dto, reconstructed=reconstructed)

    def _check_actor(self, actor_id: Any) -> None:
        if actor_id is None:
            raise E2EDtoAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.exchange" not in self._permissions:
            raise E2EDtoAuthorizationError("actor lacks permission 'e2e.exchange'")


def _exchange_event(actor_id: UUID, dto: ExpedienteDTO, at: datetime) -> Any:
    from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.dto.exchanged",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-033",
        created_at=at,
        payload={
            "api_version": dto.api_version,
            "tipo": dto.tipo,
            "estado": dto.estado,
        },
    )
