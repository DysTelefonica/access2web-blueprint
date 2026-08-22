# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# DA-2, DA-4, D90 — domain Protocols for the reset-flow service.
"""Driven ports the D90 reset-flow service depends on (DA-2, DA-4, D90).

`issue_reset_token` and `consume_reset_token` take these as keyword-only
dependencies. Real adapters (Postgres, Argon2id, mail-outbox) bind them
in PRs 43-47; the Fakes in `tests/lanzadera/auth/_fakes.py` stand in
here.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, Protocol
from uuid import UUID

if TYPE_CHECKING:
    from app.src.modules.lanzadera.domain.reset_token import ResetToken
    from app.src.modules.lanzadera.domain.user import User


class PasswordHasher(Protocol):
    """Argon2id port (DA-2, D88). Real adapter in PR 46.

    Methods are async to match the design (DA-1): the Postgres adapter
    uses ``argon2-cffi``'s async bindings under the hood, and the
    HTTP delivery that calls them already runs inside an event loop.
    The in-memory fakes in ``tests/lanzadera/auth/_fakes.py`` carry the
    matching ``async def`` shape so the contract is symmetric end to
    end.
    """

    async def hash(self, password: str) -> str: ...
    async def verify(self, password_hash: str, password: str) -> bool: ...


class UserRepository(Protocol):
    """User storage port.

    Async to match the design (DA-1). The protocol was declared sync by
    PR #42 (#288) before the design ratified async for every repository;
    this PR restates the contract to match the AsyncSession-backed
    adapters (Postgres + the in-memory Fakes) and the async delivery
    adapters that already call ``await users.<method>(...)``.

    The ``create`` / ``update_status`` / ``list_all`` operations on the
    user port are not part of the reset-flow surface (D90); they live
    on a separate Protocol (``UserAdminPort``, future) that the
    delivery/admin WU will pin. Today's admin router reaches them via
    ``getattr`` fallbacks against the in-memory fake; those are not in
    scope for the D90/W01 reset-flow surface and stay outside the
    Protocol until the dedicated WU.
    """

    async def get_by_email(self, email: str) -> User | None: ...
    async def get_by_id(self, user_id: UUID) -> User | None: ...
    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None: ...


class ResetTokenRepository(Protocol):
    """Reset-token storage port (DA-4, D90). W02 (#45) ships the Postgres adapter.

    Async to match the design (DA-1). ``insert`` takes the
    domain ``ResetToken`` value object (the in-memory fake and the
    Postgres adapter both persist it as-is; the Postgres adapter
    decomposes it into columns under the hood).
    """

    async def insert(self, token: ResetToken) -> None: ...
    async def find_unused(self, token_hash: str, now: datetime) -> ResetToken | None: ...
    async def mark_consumed(self, token_hash: str, at: datetime) -> None: ...
    async def mark_superseded(self, user_id: UUID, at: datetime) -> None: ...


class GlobalAdminRepository(Protocol):
    """Global-admin read port (D90, D91). Async."""

    async def there_is_any(self) -> bool: ...


class NotificationDelivery(Protocol):
    """Email delivery port (DA-10). Real adapter in W05 (#47).

    Async: the SMTP/SES adapter opens a network connection; a sync
    call would block the FastAPI event loop. The contract is a single
    ``send(...)`` matching the DA-10 message shape.
    """

    async def send(self, to: str, subject: str, body: str) -> None: ...


@dataclass(frozen=True)
class AuditLogEntry:
    """Audit-row value object (DA-11). Stays narrow on the auth-reset surface."""

    event_type: str
    actor_id: UUID | None
    target_id: str
    result: str
    created_at: datetime


class AuditLog(Protocol):
    """Audit emission port (DA-11). Real adapter in PR 44. Async."""

    async def append(self, event: AuditLogEntry) -> None: ...
