# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, D38, D39, D40 — pure domain module; only stdlib imports allowed.
"""Session record + lockout-policy value object (lanzadera-mvp/auth-core).

`Session` is a mutable entity: revocation shortens `expires_at` to the
current time, expiry happens automatically when the timestamp passes.
`LockoutPolicy` is a frozen value object: the threshold and duration
constants live behind a value object so the application layer can read
them from env vars at boot without leaking configuration into the rest
of the domain.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    UUID,
    dataclass,
    datetime,
)


@dataclass(frozen=True)
class LockoutPolicy:
    """Lockout policy (frozen value object).

    D38: a user is locked after `threshold` consecutive failed attempts.
    D39: the lock lasts `duration_seconds` (default 3600 s = 1 h) unless
    a global admin intervenes.

    Invariants: threshold >= 1 (a zero threshold would disable lockout
    entirely, defeating D38); duration_seconds >= 1 (a zero or negative
    duration would let the user retry immediately after a failure).
    """

    threshold: int = 5
    duration_seconds: int = 3600

    def __post_init__(self) -> None:
        if self.threshold < 1:
            raise ValueError(f"LockoutPolicy.threshold must be >= 1; got {self.threshold!r}")
        if self.duration_seconds < 1:
            raise ValueError(
                f"LockoutPolicy.duration_seconds must be >= 1; got {self.duration_seconds!r}"
            )


@dataclass
class Session:
    """User session record (mutable entity).

    Revocation is performed by setting `expires_at` to a past timestamp;
    the row stays in the table so audit consumers see the full lifecycle.
    """

    id: UUID
    user_id: UUID
    created_at: datetime
    expires_at: datetime
