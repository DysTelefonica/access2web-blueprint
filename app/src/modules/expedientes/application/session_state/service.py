"""Service implementation for EXP-CAP-045 Estado de sesión adapter (A03)."""

from __future__ import annotations

from contextvars import ContextVar
from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.session_state._evidence import (
    session_bound_event,
    session_cleared_event,
)
from app.src.modules.expedientes.application.session_state.command import (
    SessionState,
    SessionStateAuthorizationError,
    SessionStateCommand,
)


class SessionStateService:
    """Application-level adapter that binds a per-request session (CAP-045).

    The state lives in a ``contextvars.ContextVar``: the ContextVar handle is
    module-level, but its VALUE is per-execution-context (one value per asyncio
    Task, per thread, per ``contextvars.copy_context()`` block). That is the
    escape hatch documented at ``docs/05-capacidades/expedientes/access-control.md``
    line 29: no module-level mutable dict/list/set carrying per-request data,
    no global rebound across requests.

    The ``audit_log`` dependency is duck-typed (a Protocol is avoided to keep
    the cross-module ``lanzadera`` from being imported here):

        audit_log.append(event: ExpedienteAuditEvent) -> None
    """

    def __init__(self, *, audit_log: Any) -> None:
        self._audit_log = audit_log
        self._state: ContextVar[SessionState | None] = ContextVar("session_state", default=None)

    async def bind(self, command: SessionStateCommand) -> SessionState:
        actor_id = command.actor_id
        if actor_id is None or actor_id == UUID(int=0):
            raise SessionStateAuthorizationError("actor_id is required (deny-by-default)")
        new_state = SessionState(
            actor_id=actor_id,
            session_id=uuid4(),
            bound_at=datetime.now(UTC),
        )
        self._state.set(new_state)
        await self._audit_log.append(session_bound_event(actor_id, new_state.session_id))
        return new_state

    def get_current(self) -> SessionState | None:
        return self._state.get()

    async def clear(self) -> None:
        current = self._state.get()
        if current is None:
            return
        self._state.set(None)
        await self._audit_log.append(session_cleared_event(current.actor_id, current.session_id))


__all__ = ["SessionStateService"]
