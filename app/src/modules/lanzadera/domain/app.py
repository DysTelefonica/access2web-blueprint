# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-7, D52, D58, D85 — pure domain entity; only stdlib imports allowed.
"""Application catalog row (lanzadera-mvp/apps).

DA-7 + 0002: the row carries `deployment_topology` and
`requires_office_presence` derived from legacy `EjecucionEnOficina`. The
launcher fields (`Pass`, `Comando`, `URLDIrectorioIconoAplicacion`) are NOT
migrated — they belong to the runtime launch system, retired by D52.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    StrEnum,
    dataclass,
    datetime,
)


class AppTopology(StrEnum):
    """Deployment topology classification (Postgres ENUM `app_topology`).

    DA-7: `central` apps run from the platform; `office-nas` apps live on the
    office NAS and only respond when the user is inside the corporate network
    (enforced by `LocationPort` — Phase 4).
    """

    CENTRAL = "central"
    OFFICE_NAS = "office-nas"


class AppRegistrationStatus(StrEnum):
    """Lifecycle state of an app registration (Postgres ENUM `app_registration_status`).

    `pending` rows are not visible to non-admin users; `active` rows are
    visible according to user assignments; `retired` is terminal.
    """

    PENDING = "pending"
    ACTIVE = "active"
    RETIRED = "retired"


@dataclass
class App:
    """Application catalog row.

    Mutable on purpose: registration status transitions (`pending -> active ->
    retired`), name updates and `updated_at` bookkeeping.
    """

    id: int
    name: str
    short_code: str
    deployment_topology: AppTopology
    requires_office_presence: bool
    registration_status: AppRegistrationStatus
    created_at: datetime
    updated_at: datetime

    def __post_init__(self) -> None:
        if not self.name or not self.name.strip():
            raise ValueError("App.name must be a non-empty, non-whitespace string")
        if not self.short_code or not self.short_code.strip():
            raise ValueError("App.short_code must be a non-empty, non-whitespace string")
