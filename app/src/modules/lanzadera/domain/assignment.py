# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-12, H11 — pure domain entity; only stdlib imports allowed.
"""User-app-profile assignment (lanzadera-mvp/assignments).

DA-12 + H11: the assignment is the only source of truth for effective
permissions. Revocation is a soft-delete (`revoked_at`), keeping the row
visible for audit.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    UUID,
    dataclass,
    datetime,
)


@dataclass
class Assignment:
    """Effective permission triple `(user, app, profile)`.

    Revocation is non-destructive: setting `revoked_at` keeps the row in the
    table so historical reads and audit consumers see the full lifecycle.
    """

    id: UUID
    user_id: UUID
    app_id: int
    profile_id: UUID
    granted_by: UUID | None
    granted_at: datetime
    revoked_at: datetime | None = None
