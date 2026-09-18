"""Service implementation for EXP-CAP-044 CurrentPrincipal adapter (A02)."""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.current_principal._evidence import (
    principal_failed_event,
    principal_loaded_event,
)
from app.src.modules.expedientes.application.current_principal.command import (
    CurrentPrincipalAuthorizationError,
    CurrentPrincipalCommand,
    CurrentPrincipalResult,
    Principal,
)


class CurrentPrincipalService:
    """Application-level adapter that resolves the current actor's Principal (CAP-044).

    The service asks the Lanzadera container for the Principal that corresponds
    to ``command.actor_id`` and turns every failure mode (missing actor,
    Lanzadera error, empty permissions) into the same deny-by-default
    ``Principal`` that the authorization layer expects.

    The Lanzadera dependency is duck-typed (a Protocol is avoided to keep the
    cross-module ``lanzadera`` from being imported here):

        lanzadera_principal.get_principal(actor_id: UUID) -> Principal | None

    Any exception raised by the call is converted to a deny result; the
    underlying error class is recorded in the audit event so the delivery
    layer can surface the trace without leaking secrets.
    """

    def __init__(self, *, lanzadera_principal: Any, audit_log: Any) -> None:
        self._lanzadera_principal = lanzadera_principal
        self._audit_log = audit_log

    async def get_current(self, command: CurrentPrincipalCommand) -> CurrentPrincipalResult:
        actor_id = command.actor_id
        if actor_id is None or actor_id == UUID(int=0):
            raise CurrentPrincipalAuthorizationError("actor_id is required (deny-by-default)")
        try:
            retrieved = await self._lanzadera_principal.get_principal(actor_id)
        except Exception as exc:
            await self._audit_log.append(
                principal_failed_event(actor_id, "lanzadera_error", error_class=type(exc).__name__)
            )
            return CurrentPrincipalResult(principal=Principal.deny(actor_id), loaded=False)
        if retrieved is None:
            await self._audit_log.append(principal_failed_event(actor_id, "missing_principal"))
            return CurrentPrincipalResult(principal=Principal.deny(actor_id), loaded=False)
        if not retrieved.permissions:
            await self._audit_log.append(principal_failed_event(actor_id, "empty_permissions"))
            return CurrentPrincipalResult(principal=Principal.deny(actor_id), loaded=False)
        await self._audit_log.append(principal_loaded_event(actor_id, retrieved))
        return CurrentPrincipalResult(principal=retrieved, loaded=True)


__all__ = ["CurrentPrincipalService"]
