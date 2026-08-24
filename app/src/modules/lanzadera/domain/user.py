# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-3, D89 — pure domain entity; only stdlib imports allowed.
# Layer gate (scripts/check_layers.py) classifies this file as
# `app.src.modules.lanzadera.domain`; any framework import here is a violation.
"""User identity record (lanzadera-mvp/users).

DA-3: `email` is normalised to lowercase at construction time; mixed-case
input is rejected so the storage layer never has to defend against the same
invariant twice. The `name` field must be a non-empty, non-whitespace string
— empty values slip through validation in surprising places downstream.

D89 + DA-3: `password_hash` is `NULL` until `consume_reset_token` succeeds.
`dni_encrypted` carries ciphertext only; the encryption key lives in
`SecretManagerPort`, never in this module.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    UUID,
    StrEnum,
    dataclass,
    datetime,
)


class UserStatus(StrEnum):
    """Lanzadera user lifecycle states (Postgres ENUM `user_status`).

    Values match the database wire format so JSON serialisation round-trips
    without a custom encoder.
    """

    ACTIVE = "active"
    DISABLED = "disabled"
    PASSWORD_RESET_REQUIRED = "password_reset_required"
    LOCKED = "locked"


@dataclass
class User:
    """Canonical user identity record.

    Mutable on purpose: status transitions (`locked -> password_reset_required`
    after a global admin unlock, `password_reset_required -> active` after
    `consume_reset_token`), `failed_attempts` counter, and `last_login_at`
    all evolve over time. Equality is field-based; two `User`s with the
    same identity and same attribute values compare equal.
    """

    id: UUID
    email: str
    name: str
    dni_encrypted: bytes
    password_hash: str | None
    status: UserStatus
    failed_attempts: int
    last_login_at: datetime | None
    created_at: datetime
    updated_at: datetime

    def __post_init__(self) -> None:
        # DA-3: email is stored lowercase; reject any other case at the boundary.
        if self.email != self.email.lower():
            raise ValueError(f"User.email must be lowercase; got {self.email!r}")
        if not self.name or not self.name.strip():
            raise ValueError("User.name must be a non-empty, non-whitespace string")
