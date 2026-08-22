# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# In-memory port fakes for the D90 reset flow.
"""In-memory fakes for the D90 reset-flow ports (PR 42).

The async surface mirrors the protocol the Postgres adapter (W02, #45)
will pin. ``add`` / ``update_password``-style helpers stay sync because
they are *test-side convenience*, not part of the Protocol; tests
construct the user / token entities synchronously up front, then call the
async protocol methods from inside ``async def test_*`` functions.
"""

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

    async def get_by_email(self, email: str):
        return self.by_email.get(email)

    async def get_by_id(self, user_id):
        return self.by_id.get(user_id)

    async def update_password_and_activate(self, user_id, password_hash: str) -> None:
        user = self.by_id[user_id]
        user.password_hash = password_hash
        user.status = UserStatus.ACTIVE
        self.update_calls.append((user_id, password_hash))


@dataclass
class FakeResetTokenRepository:
    by_hash: dict = field(default_factory=dict)

    def add(self, token: ResetToken) -> None:
        self.by_hash[token.token_hash] = token

    async def find_unused(self, token_hash: str, now):
        row = self.by_hash.get(token_hash)
        if row is None or row.consumed_at or row.superseded_at or row.expires_at <= now:
            return None
        return row

    async def mark_consumed(self, token_hash: str, at) -> None:
        row = self.by_hash.get(token_hash)
        if row is not None:
            self.by_hash[token_hash] = dataclasses.replace(row, consumed_at=at)

    async def mark_superseded(self, user_id, at) -> None:
        for h, row in self.by_hash.items():
            if row.user_id == user_id and not row.consumed_at and not row.superseded_at:
                self.by_hash[h] = dataclasses.replace(row, superseded_at=at)

    async def insert(self, token: ResetToken) -> None:
        self.by_hash[token.token_hash] = token


@dataclass
class FakeGlobalAdminRepository:
    has_any: bool = False

    async def there_is_any(self) -> bool:
        return self.has_any


@dataclass
class FakeNotificationDelivery:
    sent: list = field(default_factory=list)

    async def send(self, to: str, subject: str, body: str) -> None:
        self.sent.append((to, subject, body))


@dataclass
class FakeAuditLog:
    entries: list = field(default_factory=list)
    next_raises: BaseException | None = None

    async def append(self, event: AuditLogEntry) -> None:
        if self.next_raises is not None:
            exc, self.next_raises = self.next_raises, None
            raise exc
        self.entries.append(event)


@dataclass
class FakePasswordHasher:
    calls: list = field(default_factory=list)
    fail_next: bool = False

    async def hash(self, password: str) -> str:
        self.calls.append(("hash", password))
        if self.fail_next:
            self.fail_next = False
            raise RuntimeError("simulated hash failure")
        return f"fake:{password}"

    async def verify(self, password: str, password_hash: str) -> bool:
        self.calls.append(("verify", password, password_hash))
        return password_hash == f"fake:{password}"
