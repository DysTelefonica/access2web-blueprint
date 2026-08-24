# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, D22, D45, D46, D110 — pure domain entity; only stdlib imports allowed.
"""Application profile row (lanzadera-mvp/profiles).

DA-12: a profile carries a JSONB-shaped `capabilities` map. Values are
`string | number | boolean` (the schema is enforced at the persistence
boundary in Phase 2 — `0003_seed_profiles`). The canonical capability set per
app is OPEN (`##ABIERTO##`, gap G-2): the seed uses `{}` until product
confirms the catalogue. Mutation here is for `active`, `name`, and the
`capabilities` map, all of which evolve with the profile.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    UUID,
    dataclass,
    datetime,
)


@dataclass
class Profile:
    """Application profile row.

    `capabilities` is a JSONB-shaped `dict[str, str | int | bool]`. The
    schema is open at the domain level (the contract lives at the persistence
    boundary); this module only enforces the bare presence of the field.
    """

    id: UUID
    app_id: int
    code: str
    name: str
    capabilities: dict[str, str | int | bool]
    active: bool
    created_at: datetime
    updated_at: datetime

    def __post_init__(self) -> None:
        if not self.code or not self.code.strip():
            raise ValueError("Profile.code must be a non-empty, non-whitespace string")
        if not self.name or not self.name.strip():
            raise ValueError("Profile.name must be a non-empty, non-whitespace string")
