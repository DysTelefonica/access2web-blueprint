# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Idempotent seed of the 8 profile × 8 app matrix (DA-12, 0003_seed_profiles).

This module is the canonical implementation of the ``0003_seed_profiles``
migration described in ``openspec/changes/lanzadera-mvp/specs/profiles/spec.md``.

The seed is idempotent: re-running it is a no-op because each
``(app_id, code)`` pair is unique-constrained in the database and
the repository's ``get_by_code`` returns the existing row.

Run:
    python -m app.src.modules.lanzadera.application.seed_profiles

Or via the CLI once ``gentle-ai platform seed-profiles`` is wired in.

Static catalogue lives in sibling modules (W-TEST split):
  ``seed_profile_apps``        — 8 legacy apps from TbAplicaciones
  ``seed_profile_codes``       — 8 profile codes + display names
  ``seed_capability_*.py``     — one module per 2-3 profile capability maps
  ``seed_profiles``            — this file: runner only

Profile codes (from ``legacy_role_map.LEGACY_ROLE_MAP``):
    DEFAULT        — fallback when no explicit profile matches
    ADMIN          — maps to legacy EsAdministrador
    CALIDAD        — maps to legacy EsCalidad
    CALIDAD_AVISOS — maps to legacy EsCalidadAvisos
    TECNICO        — maps to legacy EsTecnico
    ECONOMIA       — maps to legacy EsEconomía
    SECRETARIA     — maps to legacy EsSecretaría
    SIN_ACCESO     — exclusive: blocks all access (DA-12 cortocircuito)
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

from app.src.modules.lanzadera.domain.audit_event import AuditEvent
from app.src.modules.lanzadera.domain.ports import AuditLog
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)

# Data modules — each ≤ 100 mutation sites (W-TEST refactor).
from app.src.modules.lanzadera.application.seed_profile_apps import (
    APP_CATALOGUE as _APP_CATALOGUE,
)
from app.src.modules.lanzadera.application.seed_profile_codes import (
    PROFILE_CODES as _PROFILE_CODES,
)
from app.src.modules.lanzadera.application.seed_capability_DEFAULT_ADMIN_CALIDAD import (
    DEFAULT as _DEFAULT,
    ADMIN as _ADMIN,
    CALIDAD as _CALIDAD,
)
from app.src.modules.lanzadera.application.seed_capability_CALIDAD_AVISOS_TECNICO_ECONOMIA import (
    CALIDAD_AVISOS as _CALIDAD_AVISOS,
    TECNICO as _TECNICO,
    ECONOMIA as _ECONOMIA,
)
from app.src.modules.lanzadera.application.seed_capability_SECRETARIA_SIN_ACCESO import (
    SECRETARIA as _SECRETARIA,
    SIN_ACCESO as _SIN_ACCESO,
)

#: Lookup table from profile code to its 6-flag capability map.
#: Built at module import time from the capability data modules.
_CAPABILITIES_BY_CODE: dict[str, dict[str, bool]] = {
    "DEFAULT": _DEFAULT,
    "ADMIN": _ADMIN,
    "CALIDAD": _CALIDAD,
    "CALIDAD_AVISOS": _CALIDAD_AVISOS,
    "TECNICO": _TECNICO,
    "ECONOMIA": _ECONOMIA,
    "SECRETARIA": _SECRETARIA,
    "SIN_ACCESO": _SIN_ACCESO,
}


async def seed_profiles(
    profiles: ProfileRepositoryPort,
    audit: AuditLog,
    *,
    now: datetime | None = None,
) -> int:
    """Insert every (app, profile) row if it does not already exist.

    Returns the number of rows created. Idempotent: rows that already
    exist are left untouched.
    """
    if now is None:
        now = datetime.now(UTC)

    created = 0
    for app in _APP_CATALOGUE:
        app_id: int = app["id"]  # type: ignore[assignment]
        for pdef in _PROFILE_CODES:
            existing = await profiles.get_by_code(app_id, pdef["code"])
            if existing is not None:
                continue  # already seeded — skip

            from app.src.modules.lanzadera.domain.profile import Profile

            profile = Profile(
                id=uuid4(),
                app_id=app_id,
                code=pdef["code"],
                name=pdef["name"],
                capabilities=_CAPABILITIES_BY_CODE[pdef["code"]].copy(),
                active=True,
                created_at=now,
                updated_at=now,
            )
            await profiles.create(profile)
            created += 1

    if created > 0:
        await audit.append(
            AuditEvent(  # type: ignore[arg-type]
                id=uuid4(),
                event_type="profiles.seed",
                actor_id=None,
                target_id="profiles.seed",
                module="lanzadera",
                result="success",
                correlation_id=uuid4(),
                payload={"created": created},
                created_at=now,
            )
        )

    return created
