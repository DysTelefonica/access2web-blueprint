# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, D21, D42 — pure domain value object; only stdlib imports allowed.
"""Global-admin row (lanzadera-mvp/global_admins).

D21: the global-admin role is materialised as a row in `global_admins`,
NOT as a column on `users`. `GlobalAdmin` is therefore a one-attribute
value object — `user_id` — and frozen by design (revoke / grant live in
the application layer and produce delete / insert operations).
"""

from __future__ import annotations

from dataclasses import dataclass
from uuid import UUID


@dataclass(frozen=True)
class GlobalAdmin:
    """Global-admin membership record (frozen value object).

    The presence of this row is the only signal the auth layer trusts.
    """

    user_id: UUID
