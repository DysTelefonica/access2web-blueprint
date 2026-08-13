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
    """Argon2id port (DA-2, D88). Real adapter in PR 46."""

    def hash(self, password: str) -> str: ...
    def verify(self, password_hash: str, password: str) -> bool: ...


class UserRepository(Protocol):
    """User storage port. PR 43 ships the full contract."""

    def get_by_email(self, email: str) -> User | None: ...
    def get_by_id(self, user_id: UUID) -> User | None: ...
    def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None: ...


class ResetTokenRepository(Protocol):
    """Reset-token storage port (DA-4, D90). PR 45 ships the Postgres adapter."""

    def insert(self, token: ResetToken) -> None: ...
    def find_unused(self, token_hash: str, now: datetime) -> ResetToken | None: ...
    def mark_consumed(self, token_hash: str, at: datetime) -> None: ...
    def mark_superseded(self, user_id: UUID, at: datetime) -> None: ...


class GlobalAdminRepository(Protocol):
    """Global-admin read port (D90, D91)."""

    def there_is_any(self) -> bool: ...


class NotificationDelivery(Protocol):
    """Email delivery port (DA-10). Real adapter in PR 47."""

    def send(self, to: str, subject: str, body: str) -> None: ...


@dataclass(frozen=True)
class AuditLogEntry:
    """Audit-row value object (DA-11). Stays narrow on the auth-reset surface."""

    event_type: str
    actor_id: UUID | None
    target_id: str
    result: str
    created_at: datetime


class AuditLog(Protocol):
    """Audit emission port (DA-11). Real adapter in PR 44."""

    def append(self, event: AuditLogEntry) -> None: ...
