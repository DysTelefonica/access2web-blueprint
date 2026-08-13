# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# In-memory test fakes for the D90 reset-flow ports.
"""In-memory fakes for the D90 reset-flow ports (PR 42)."""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from app.src.modules.lanzadera.domain.user import User, UserStatus

if TYPE_CHECKING:
    from app.src.modules.lanzadera.domain.ports import AuditLogEntry
    from app.src.modules.lanzadera.domain.reset_token import ResetToken


@dataclass
class FakeUserRepository:
    by_id: dict = field(default_factory=dict)
    by_email: dict = field(default_factory=dict)
    update_calls: list = field(default_factory=list)

    def add(self, user: User) -> None:
        self.by_id[user.id] = user
        self.by_email[user.email] = user

    def get_by_email(self, email: str) -> User | None:
        return self.by_email.get(email)

    def get_by_id(self, user_id) -> User | None:
        return self.by_id.get(user_id)

    def update_password_and_activate(self, user_id, password_hash: str) -> None:
        user = self.by_id[user_id]
        user.password_hash = password_hash
        user.status = UserStatus.ACTIVE
        self.update_calls.append((user_id, password_hash))


@dataclass
class FakeResetTokenRepository:
    by_hash: dict = field(default_factory=dict)

    def add(self, token: "ResetToken") -> None:
        self.by_hash[token.token_hash] = token

    def find_unused(self, token_hash: str, now) -> "ResetToken | None":
        token = self.by_hash.get(token_hash)
        if token is None:
            return None
        if token.consumed_at is not None or token.superseded_at is not None:
            return None
        if token.expires_at <= now:
            return None
        return token

    def mark_consumed(self, token_hash: str, at) -> None:
        token = self.by_hash.get(token_hash)
        if token is not None:
            self.by_hash[token_hash] = dataclasses.replace(token, consumed_at=at)

    def mark_superseded(self, user_id, at) -> None:
        for h, token in self.by_hash.items():
            if token.user_id == user_id and token.consumed_at is None and token.superseded_at is None:
                self.by_hash[h] = dataclasses.replace(token, superseded_at=at)

    def insert(self, token: "ResetToken") -> None:
        self.by_hash[token.token_hash] = token


@dataclass
class FakeGlobalAdminRepository:
    has_any: bool = False

    def there_is_any(self) -> bool:
        return self.has_any


@dataclass
class FakeNotificationDelivery:
    sent: list = field(default_factory=list)

    def send(self, to: str, subject: str, body: str) -> None:
        self.sent.append((to, subject, body))


@dataclass
class FakeAuditLog:
    entries: list = field(default_factory=list)
    next_raises: BaseException | None = None

    def append(self, event: "AuditLogEntry") -> None:
        if self.next_raises is not None:
            exc, self.next_raises = self.next_raises, None
            raise exc
        self.entries.append(event)


@dataclass
class FakePasswordHasher:
    calls: list = field(default_factory=list)
    fail_next: bool = False

    def hash(self, password: str) -> str:
        self.calls.append(("hash", password))
        if self.fail_next:
            self.fail_next = False
            raise RuntimeError("simulated hash failure")
        return f"fake:{password}"

    def verify(self, password_hash: str, password: str) -> bool:
        self.calls.append(("verify", password_hash, password))
        return password_hash == f"fake:{password}"
