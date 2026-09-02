# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#540)
"""Admin use case: logout (D-W62-2, DA-11).

Revokes a single ``Session`` by shortening its ``expires_at`` to the
current ``now``. The caller (delivery, PR-5) resolves the session id
from the JWT Bearer token before invoking this use case.

D-W62-2: revocation keeps the row (DA-11 audit-consumer visibility) and
shortens ``expires_at`` instead of deleting; the auth middleware (PR-5)
rejects requests whose ``expires_at`` is in the past.

DA-11: the audit row is appended inside the same logical operation
so the security log captures the session termination alongside the
creation row from the login use case.

D-W62-4: logout revokes the session presented in the request — not
``all_sessions_for_user``. The "logout everywhere" slice (revoke every
session for the user) is a separate epic when the use case arises.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import SessionNotFoundError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
)
from app.src.modules.lanzadera.domain.ports.session_repository import (
    SessionRepository,
)


async def logout(
    session_id: UUID,
    *,
    now: datetime,
    sessions: SessionRepository,
    audit: AuditLog,
    actor_id: UUID | None = None,
) -> None:
    """Revoke the session identified by ``session_id``.

    Sequence (DA-11, D-W62-2):
        1. ``SessionRepository.get_by_id(session_id)`` — fetch the row.
        2. ``SessionRepository.revoke(session_id, now)`` — shorten
           ``expires_at`` to ``now`` (idempotent on the second call;
           ``expires_at`` stays at the earlier of the two timestamps).
        3. ``AuditLog.append(...)`` — security log row with the
           terminated session id as ``target_id``.

    Raises:
        ``SessionNotFoundError`` — no row exists for ``session_id``.
        The caller (delivery, PR-5) maps this to HTTP 401 to avoid
        leaking which sessions were once valid; the audit row records
        the attempt with ``target_id=session_id``.

    Idempotency:
        Revoking an already-revoked session succeeds without raising —
        the underlying ``SessionRepository.revoke`` is idempotent and
        ``expires_at`` is shortened to the earlier of ``now`` and the
        current ``expires_at``. ``SessionNotFoundError`` is only raised
        when the row does not exist at all.

    Side effects (DA-11, in order):
        1. ``SessionRepository.get_by_id(session_id)`` — read.
        2. ``SessionRepository.revoke(session_id, now)`` — write.
        3. ``AuditLog.append(...)`` — security log row.
    """
    if session_id is None:
        raise ValueError("logout.session_id must be a UUID")

    existing = await sessions.get_by_id(session_id)
    if existing is None:
        await audit.append(
            AuditLogEntry(
                event_type="auth.logout.failure",
                actor_id=actor_id,
                target_id=str(session_id),
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={"reason": "unknown_session"},
            )
        )
        raise SessionNotFoundError(f"session {session_id!r} not found")

    revoked = await sessions.revoke(session_id, now)
    if not revoked:
        # Defensive: the row existed at ``get_by_id`` but disappeared
        # between the two reads. Treat as not-found so the caller maps
        # the response identically (D-W62-2: no enumeration via timing).
        await audit.append(
            AuditLogEntry(
                event_type="auth.logout.failure",
                actor_id=actor_id,
                target_id=str(session_id),
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={"reason": "session_disappeared"},
            )
        )
        raise SessionNotFoundError(f"session {session_id!r} not found")

    await audit.append(
        AuditLogEntry(
            event_type="auth.logout.success",
            actor_id=actor_id,
            target_id=str(session_id),
            result="success",
            created_at=now,
            module="lanzadera",
            correlation_id=uuid4(),
            payload={"user_id": str(existing.user_id)},
        )
    )


__all__ = ["logout"]
