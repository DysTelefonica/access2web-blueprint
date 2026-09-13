"""Service implementation for EXP-CAP-041 E2E session (E07)."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.e2e_session._evidence import (
    session_change,
    session_event,
)
from app.src.modules.expedientes.application.e2e_session.command import (
    E2ESessionAuthorizationError,
    E2ESessionError,
    E2ESessionResult,
    E2ESessionStatus,
    E2ESessionValidationError,
)


class E2ESessionService:
    """Application-level E2E session lifecycle (CAP-041).

    The service opens a new session, resumes or closes an existing
    one, and emits one audit event per state transition. The session
    store is injected so production wiring supplies the persistent
    implementation while tests use an in-memory map.

    Session ownership is enforced on resume: a session is bound to
    the actor that opened it. A different actor cannot resume it.
    """

    def __init__(
        self,
        *,
        audit_log: Any,
        store: Any,
        permissions: set[str],
    ) -> None:
        self._audit_log = audit_log
        self._store = store
        self._permissions = permissions

    def grant(self, permission: str) -> None:
        self._permissions.add(permission)

    async def open(self, *, actor_id: UUID | None) -> E2ESessionResult:
        self._check_actor(actor_id)
        assert actor_id is not None
        now = datetime.now(UTC)
        session: dict[str, Any] = {
            "session_id": uuid4(),
            "actor_id": actor_id,
            "status": E2ESessionStatus.ACTIVE.value,
            "opened_at": now,
            "last_seen_at": now,
        }
        try:
            await self._store.create(session)
        except Exception as exc:
            raise E2ESessionError(f"session failed for actor {actor_id!r}: {exc}") from exc
        await self._audit_log.append(
            session_event(actor_id, "e2e.session.opened", session["session_id"])
        )
        return _result(session)

    async def resume(self, *, actor_id: UUID | None, session_id: UUID) -> E2ESessionResult:
        self._check_actor(actor_id)
        assert actor_id is not None
        session = await self._load(session_id)
        if session["actor_id"] != actor_id:
            raise E2ESessionAuthorizationError(
                f"session {session_id!r} does not belong to actor {actor_id!r}"
            )
        if session["status"] == E2ESessionStatus.CLOSED.value:
            raise E2ESessionValidationError(f"session {session_id!r} is closed")
        session["last_seen_at"] = datetime.now(UTC)
        try:
            await self._store.update(session)
        except Exception as exc:
            raise E2ESessionError(f"session failed for actor {actor_id!r}: {exc}") from exc
        await self._audit_log.append(session_event(actor_id, "e2e.session.resumed", session_id))
        return _result(session)

    async def close(self, *, actor_id: UUID | None, session_id: UUID) -> E2ESessionResult:
        self._check_actor(actor_id)
        assert actor_id is not None
        session = await self._load(session_id)
        if session["actor_id"] != actor_id:
            raise E2ESessionAuthorizationError(
                f"session {session_id!r} does not belong to actor {actor_id!r}"
            )
        session["status"] = E2ESessionStatus.CLOSED.value
        session["last_seen_at"] = datetime.now(UTC)
        try:
            await self._store.update(session)
        except Exception as exc:
            raise E2ESessionError(f"session failed for actor {actor_id!r}: {exc}") from exc
        await self._audit_log.append(session_event(actor_id, "e2e.session.closed", session_id))
        await self._audit_log.append(
            session_change(actor_id, session_id, E2ESessionStatus.CLOSED.value)
        )
        return _result(session)

    async def _load(self, session_id: UUID) -> dict[str, Any]:
        try:
            session: dict[str, Any] | None = await self._store.get(session_id)
        except Exception as exc:
            raise E2ESessionError(f"session failed for {session_id!r}: {exc}") from exc
        if session is None:
            raise E2ESessionValidationError(f"session {session_id!r} not found")
        return session

    def _check_actor(self, actor_id: UUID | None) -> None:
        if actor_id is None:
            raise E2ESessionAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.session" not in self._permissions:
            raise E2ESessionAuthorizationError("actor lacks permission 'e2e.session'")


def _result(session: dict[str, Any]) -> E2ESessionResult:
    return E2ESessionResult(
        session_id=session["session_id"],
        actor_id=session["actor_id"],
        status=E2ESessionStatus(session["status"]),
        opened_at=session["opened_at"],
        last_seen_at=session["last_seen_at"],
    )


__all__ = ["E2ESessionService"]
