"""SessionRepositoryPort — D-W62-2, DA-1.

The persistence boundary for the W62 auth flow's session rows.
``login`` (PR-2) calls ``create`` after a successful password verification;
the auth middleware (PR-5) calls ``get_by_id`` on every authenticated
request; ``logout`` (PR-3) calls ``revoke``. Revocation shortens
``expires_at`` to a past timestamp and keeps the row — the audit
consumer (DA-11) reads the full lifecycle, not just the live sessions.

DA-1: every method is async because the Postgres adapter uses
``AsyncSession``; the in-memory ``FakeSessionRepository`` in
``tests/lanzadera/_fakes.py`` mirrors the async shape so the contract is
symmetric end to end.
"""

from __future__ import annotations

from datetime import datetime

from app.src.modules.lanzadera.domain._imports import UUID
from app.src.modules.lanzadera.domain.ports._imports import Protocol
from app.src.modules.lanzadera.domain.session import Session


class SessionRepositoryPort(Protocol):
    """Persistence boundary for ``Session`` rows."""

    async def create(self, session: Session) -> Session:
        """Persist a fresh session row.

        Returns the canonical ``Session`` (the Postgres adapter may stamp
        a server-side ``created_at`` if the caller omitted it).
        """
        ...

    async def get_by_id(self, session_id: UUID) -> Session | None:
        """Return the row iff it exists; ``None`` otherwise.

        The caller (auth middleware) is responsible for the expiry
        check — the repository does not silently filter expired rows,
        so the audit consumer can still see them.
        """
        ...

    async def revoke(self, session_id: UUID, at: datetime) -> bool:
        """Shorten ``expires_at`` to ``at``. Idempotent on the second call.

        Returns ``True`` if the row was found and updated, ``False`` if
        no such session exists (caller can decide whether that is a
        programming error or an expected idempotent retry).
        """
        ...


SessionRepository = SessionRepositoryPort


__all__ = ["SessionRepository", "SessionRepositoryPort"]
