"""In-memory port fakes for the AD2 repository tests (PR 45).

Each fake satisfies the corresponding Port contract structurally (duck-typed
in Python, ``typing.Protocol``-compatible by shape) so the same fake is
usable both in ``tests/lanzadera/<domain>/`` and in the application
layer tests once those WUs land.

Why in-memory: the four Pg adapters in ``app/src/modules/lanzadera/adapters/repos/``
are still skeletons (the underlying migrations ``0003_seed_assignments``
+ ``0006_seed_audit`` haven't landed). Without a real DB we pin the
contract today; when the bodies arrive the same test suite exercises
the real adapter — the protocol shapes don't change.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

if TYPE_CHECKING:
    from app.src.modules.lanzadera.domain.assignment import Assignment
    from app.src.modules.lanzadera.domain.audit_event import AuditEvent
    from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin
    from app.src.modules.lanzadera.domain.reset_token import ResetToken


# ---------------------------------------------------------------------------
# Assignment
# ---------------------------------------------------------------------------


@dataclass
class FakeAssignmentRepository:
    """In-memory store keyed by ``(user_id, app_id, profile_id)``.

    Tracks every mutation in :attr:`calls` so contract tests can assert
    that the adapter issued exactly one INSERT / UPDATE per case-of-use.
    """

    by_id: dict[UUID, Assignment] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    def add(self, assignment: Assignment) -> None:
        self.by_id[assignment.id] = assignment
        self.calls.append("add")

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        from app.src.modules.lanzadera.domain.assignment import Assignment

        self.calls.append(f"create({user_id},{app_id},{profile_id})")
        assignment = Assignment(
            id=uuid4(),
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            granted_by=None,
            granted_at=datetime.now(UTC),
            revoked_at=None,
        )
        self.by_id[assignment.id] = assignment
        return assignment

    async def list_for_user(self, user_id: UUID) -> list[Assignment]:
        self.calls.append(f"list_for_user({user_id})")
        return [a for a in self.by_id.values() if a.user_id == user_id and a.revoked_at is None]

    async def list_for_app(self, app_id: int) -> list[Assignment]:
        self.calls.append(f"list_for_app({app_id})")
        return [a for a in self.by_id.values() if a.app_id == app_id and a.revoked_at is None]

    async def effective_permissions(self, user_id: UUID, app_id: int) -> list[str]:
        self.calls.append(f"effective_permissions({user_id},{app_id})")
        # The contract returns capabilities names from the profile — but the
        # profile is owned by a different port. The fake returns an empty
        # list; tests that exercise the application layer's profile lookup
        # can layer a real Profile fake on top.
        return []


# ---------------------------------------------------------------------------
# Reset token
# ---------------------------------------------------------------------------


@dataclass
class FakeResetTokenRepository:
    """In-memory store keyed by ``token_hash``.

    Mirrors the production semantics: tokens are looked up by hash, marked
    consumed at most once, and superseded on the user level.
    """

    by_hash: dict[str, ResetToken] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    def add(self, token: ResetToken) -> None:
        self.by_hash[token.token_hash] = token
        self.calls.append("add")

    async def insert(self, user_id: UUID, token_hash: str, expires_at: datetime) -> ResetToken:
        from app.src.modules.lanzadera.domain.reset_token import ResetToken

        self.calls.append(f"insert({user_id},{token_hash[:8]}...)")
        token = ResetToken(
            id=uuid4(),
            user_id=user_id,
            token_hash=token_hash,
            expires_at=expires_at,
            consumed_at=None,
            superseded_at=None,
            created_at=datetime.now(UTC),
        )
        self.by_hash[token_hash] = token
        return token

    async def find_unused(self, token_hash: str) -> ResetToken | None:
        self.calls.append(f"find_unused({token_hash[:8]}...)")
        token = self.by_hash.get(token_hash)
        if token is None:
            return None
        if token.consumed_at is not None or token.superseded_at is not None:
            return None
        if token.expires_at < datetime.now(UTC):
            return None
        return token

    async def mark_consumed(self, token_hash: str, at: datetime) -> None:
        import dataclasses

        self.calls.append(f"mark_consumed({token_hash[:8]}...)")
        token = self.by_hash.get(token_hash)
        if token is not None and token.consumed_at is None:
            # ``ResetToken`` is frozen — produce a new instance via
            # ``dataclasses.replace`` so the fake does not mutate it.
            self.by_hash[token_hash] = dataclasses.replace(token, consumed_at=at)

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        import dataclasses

        self.calls.append(f"mark_superseded({user_id})")
        for token in list(self.by_hash.values()):
            if token.user_id == user_id and token.superseded_at is None:
                self.by_hash[token.token_hash] = dataclasses.replace(token, superseded_at=at)

    async def purge_expired(self, now: datetime) -> int:
        self.calls.append("purge_expired")
        # An expired token is one whose ``expires_at`` is in the past
        # RELATIVE to ``now``. Because the domain ``ResetToken`` rejects
        # ``expires_at <= created_at``, we cannot construct a fully-expired
        # token via the normal ``insert`` path. Tests that need an expired
        # token either pre-load it via :meth:`add` or use
        # ``dataclasses.replace`` to backdate ``created_at`` so the
        # ``expires_at`` constraint still holds — see the contract tests
        # for an example.
        expired = [h for h, t in self.by_hash.items() if t.expires_at < now]
        for h in expired:
            del self.by_hash[h]
        return len(expired)


# ---------------------------------------------------------------------------
# Global admin
# ---------------------------------------------------------------------------


@dataclass
class FakeGlobalAdminRepository:
    """In-memory store of ``user_id`` memberships."""

    members: set[UUID] = field(default_factory=set)
    calls: list[str] = field(default_factory=list)

    async def list_all(self) -> list[GlobalAdmin]:
        from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin

        self.calls.append("list_all")
        return [GlobalAdmin(user_id=u) for u in self.members]

    async def is_global_admin(self, user_id: UUID) -> bool:
        self.calls.append(f"is_global_admin({user_id})")
        return user_id in self.members

    async def grant(self, user_id: UUID) -> None:
        self.calls.append(f"grant({user_id})")
        if user_id in self.members:
            raise ValueError(f"user {user_id} is already a global admin")
        self.members.add(user_id)

    async def revoke(self, user_id: UUID) -> None:
        self.calls.append(f"revoke({user_id})")
        if user_id not in self.members:
            # Idempotent semantics: revoke of a non-admin is a no-op (not a
            # hard error). Production policy differs (the adapter raises),
            # but this fake matches the lenient path so revoke tests can
            # exercise both branches without plumbing fake state per branch.
            return
        if len(self.members) == 1:
            raise ValueError("cannot revoke the last global admin")
        self.members.discard(user_id)


# ---------------------------------------------------------------------------
# Audit log
# ---------------------------------------------------------------------------


@dataclass
class FakeAuditLogPort:
    """Append-only in-memory audit log (with a fan-out list API)."""

    events: list[AuditEvent] = field(default_factory=list)
    calls: list[str] = field(default_factory=list)

    async def append(self, event: AuditEvent) -> None:
        self.calls.append(f"append({event.event_type})")
        self.events.append(event)

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> list[AuditEvent]:
        self.calls.append(f"list_for_actor({actor_id},{since})")
        return [e for e in self.events if e.actor_id == actor_id and e.created_at >= since]
