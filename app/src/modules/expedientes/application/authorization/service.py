"""Service implementation for EXP-CAP-043 authorization policy (A01)."""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.authorization._evidence import (
    authz_event,
)
from app.src.modules.expedientes.application.authorization.command import (
    EFFECTIVE_DENY,
    AuthorizationAuthorizationError,
    AuthorizationCommand,
    AuthorizationResult,
    AuthorizationValidationError,
)


class AuthorizationService:
    """Application-level deny-by-default authorization policy (CAP-043).

    The service is intentionally tiny: a capability is granted when
    the actor's effective permissions set contains it; otherwise the
    service returns ``allowed=False`` with ``denied_reason="deny_by_default"``.
    Every decision emits one audit event so the delivery layer can
    surface the trace.
    """

    def __init__(self, *, audit_log: Any) -> None:
        self._audit_log = audit_log

    async def authorize(self, command: AuthorizationCommand) -> AuthorizationResult:
        if command.actor_id is None or command.actor_id == UUID(int=0):
            raise AuthorizationAuthorizationError("actor_id is required (deny-by-default)")
        capability = command.capability.strip() if command.capability else ""
        if not capability:
            raise AuthorizationValidationError("capability is required (deny-by-default)")
        allowed = capability in command.permissions
        result = AuthorizationResult(
            actor_id=command.actor_id,
            capability=capability,
            allowed=allowed,
            denied_reason=None if allowed else EFFECTIVE_DENY,
        )
        await self._audit_log.append(authz_event(command.actor_id, result))
        return result


__all__ = ["AuthorizationService"]
