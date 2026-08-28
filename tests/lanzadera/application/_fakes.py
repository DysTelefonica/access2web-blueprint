# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""In-memory port fakes for the admin use cases (W54, issue #43).

Extends the auth-reset fakes with the additional ports the admin use
cases need. ``FakeUserRepository`` is re-exported from the auth
package; new fakes here cover GlobalAdminRepository, AppRepository,
ProfileRepository, AssignmentRepository, SecretManager, and
CredentialHasher.
"""

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.app import App
from app.src.modules.lanzadera.domain.assignment import Assignment
from app.src.modules.lanzadera.domain.profile import Profile
from tests.lanzadera.auth._fakes import (
    FakeAuditLog,
    FakeUserRepository,
)

if TYPE_CHECKING:
    pass


@dataclass
class FakeGlobalAdminRepository:
    rows: dict = field(default_factory=dict)  # user_id -> GlobalAdmin

    async def there_is_any(self) -> bool:
        return bool(self.rows)

    async def is_global_admin(self, user_id: UUID) -> bool:
        return user_id in self.rows

    async def list_all(self) -> Sequence:
        from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin

        return [GlobalAdmin(user_id=uid) for uid in self.rows]

    async def grant(self, user_id: UUID) -> None:
        from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin

        self.rows[user_id] = GlobalAdmin(user_id=user_id)

    async def revoke(self, user_id: UUID) -> None:
        if user_id in self.rows and len(self.rows) <= 1:
            raise ValueError(
                f"cannot revoke global admin {user_id!r}: would leave the system with zero admins"
            )
        self.rows.pop(user_id, None)


@dataclass
class FakeAppRepository:
    by_id: dict = field(default_factory=dict)

    def add(self, app: App) -> None:
        self.by_id[app.id] = app

    async def get_by_id(self, app_id: int):
        return self.by_id.get(app_id)

    async def list_active(self) -> Sequence[App]:
        return [a for a in self.by_id.values() if a.registration_status.value == "active"]


@dataclass
class FakeProfileRepository:
    by_code: dict = field(default_factory=dict)  # (app_id, code) -> Profile
    by_id: dict = field(default_factory=dict)  # profile_id -> Profile

    def add(self, profile: Profile) -> None:
        self.by_code[(profile.app_id, profile.code)] = profile
        self.by_id[profile.id] = profile

    async def get_by_code(self, app_id: int, code: str):
        return self.by_code.get((app_id, code))

    async def get_by_id(self, profile_id: UUID):
        return self.by_id.get(profile_id)

    async def list_for_app(self, app_id: int) -> Sequence[Profile]:
        return [p for (a, _c), p in self.by_code.items() if a == app_id]


@dataclass
class FakeAssignmentRepository:
    by_user: dict = field(default_factory=dict)  # user_id -> list[Assignment]
    by_pair: dict = field(default_factory=dict)  # (user_id, app_id) -> Assignment

    def add(self, assignment: Assignment) -> None:
        self.by_user.setdefault(assignment.user_id, []).append(assignment)
        self.by_pair[(assignment.user_id, assignment.app_id)] = assignment

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        if (user_id, app_id) in self.by_pair:
            raise ValueError(f"duplicate assignment for user={user_id} app={app_id}")
        assignment = Assignment(
            id=uuid4(),
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            granted_by=None,
            granted_at=datetime.now(),
            revoked_at=None,
        )
        self.add(assignment)
        return assignment

    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]:
        return [a for a in self.by_user.get(user_id, []) if a.revoked_at is None]

    async def list_for_app(self, app_id: int) -> Sequence[Assignment]:
        return [a for a in self.by_pair.values() if a.app_id == app_id and a.revoked_at is None]

    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]:
        return []


@dataclass
class FakeSecretManager:
    encrypted: list = field(default_factory=list)

    def encrypt(self, plaintext: str) -> bytes:
        ciphertext = plaintext.encode("utf-8")[::-1]
        self.encrypted.append(plaintext)
        return ciphertext

    def decrypt(self, ciphertext: bytes) -> str:
        return ciphertext[::-1].decode("utf-8")


@dataclass
class FakeCredentialHasher:
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


__all__ = [
    "FakeAppRepository",
    "FakeAssignmentRepository",
    "FakeAuditLog",
    "FakeCredentialHasher",
    "FakeGlobalAdminRepository",
    "FakeProfileRepository",
    "FakeSecretManager",
    "FakeUserRepository",
]
