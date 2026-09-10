# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp
"""Seed profiles for the 8 known legacy apps (DA-12, 0003_seed_profiles).

This module is the canonical implementation of the `0003_seed_profiles`
migration described in `openspec/changes/lanzadera-mvp/specs/profiles/spec.md`.

The seed is idempotent: re-running it is a no-op because each
``(app_id, code)`` pair is unique-constrained in the database and
the repository's ``get_by_code`` returns the existing row.

Run:
    python -m app.src.modules.lanzadera.application.seed_profiles

Or via the CLI once `gentle-ai platform seed-profiles` is wired in.

Profile codes (from `legacy_role_map.LEGACY_ROLE_MAP`):
    DEFAULT        — fallback when no explicit profile matches
    ADMIN          — maps to legacy EsAdministrador
    CALIDAD        — maps to legacy EsCalidad
    CALIDAD_AVISOS — maps to legacy EsCalidadAvisos
    TECNICO        — maps to legacy EsTecnico
    ECONOMIA       — maps to legacy EsEconomía
    SECRETARIA     — maps to legacy EsSecretaría
    SIN_ACCESO     — exclusive: blocks all access (DA-12 cortocircuito)

Capability maps are intentionally empty (``{}``) per the spec:
    > The canonical capability set per app is OPEN — the seed uses {}
    > until product confirms the catalogue.
The JSONB contract accepts any ``str | int | bool`` values.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.ports import AuditLog, AuditLogEntry
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)


# ---------------------------------------------------------------------------
# Static catalogue — eight apps from the legacy `TbAplicaciones` baseline.
# IDs are confirmed in docs/06-autorizacion-legacy-matriz.md
# (Verified-runtime-schema/aggregate).
# ---------------------------------------------------------------------------

_APP_CATALOGUE: list[dict] = [
    {"id": 5,  "name": "Gestion_Riesgos",   "short_code": "RIESGOS"},
    {"id": 6,  "name": "Brass",             "short_code": "BRASS"},
    {"id": 8,  "name": "No_Conformidades",  "short_code": "NOCONF"},
    {"id": 12, "name": "Lanzadera",         "short_code": "LANZ"},
    {"id": 17, "name": "HPS",               "short_code": "HPS"},
    {"id": 19, "name": "Expedientes",       "short_code": "EXP"},
    {"id": 22, "name": "HPS_Solicitudes",   "short_code": "HPSSOL"},
    {"id": 23, "name": "Condor",            "short_code": "CONDOR"},
]

# Profile codes and their display names.
_PROFILE_CODES: list[dict] = [
    {"code": "DEFAULT",        "name": "Usuario por defecto"},
    {"code": "ADMIN",         "name": "Administrador"},
    {"code": "CALIDAD",       "name": "Calidad"},
    {"code": "CALIDAD_AVISOS","name": "Calidad + Avisos"},
    {"code": "TECNICO",       "name": "Técnico"},
    {"code": "ECONOMIA",      "name": "Economía"},
    {"code": "SECRETARIA",    "name": "Secretaría"},
    {"code": "SIN_ACCESO",    "name": "Sin acceso"},
]


# ---------------------------------------------------------------------------
# Seed logic
# ---------------------------------------------------------------------------

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
        app_id: int = app["id"]
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
                capabilities={},  # intentionally empty — see module docstring
                active=True,
                created_at=now,
                updated_at=now,
            )
            await profiles.create(profile)
            created += 1

    if created > 0:
        await audit.append(
            AuditLogEntry(
                event_type="profiles.seed",
                actor_id=None,
                target_id=None,
                module="lanzadera",
                result="success",
                correlation_id=uuid4(),
                payload={"created": created},
                created_at=now,
            )
        )

    return created


# ---------------------------------------------------------------------------
# Standalone runner — invoke directly with Python.
# ---------------------------------------------------------------------------

async def _main() -> None:
    """Build a real container and run the seed.

    Requires the same environment variables as the running application:
        DATABASE_URL   — Postgres connection string
        JWT_SECRET     — for JWT operations
        SECRET_KEY     — for credential encryption

    Exits with code 0 on success, 1 on error.
    """
    import os

    from app.src.modules.lanzadera.di.container import LanzaderaContainer
    from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
        AsyncSessionFactory,
    )

    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        print("ERROR: DATABASE_URL environment variable is not set.", file=__import__("sys").stderr)
        raise SystemExit(1)

    print("Building container...")
    session_factory = AsyncSessionFactory(database_url)
    container = LanzaderaContainer(session_factory=session_factory)

    print("Seeding profiles...")
    created = await seed_profiles(
        profiles=container.profile_repo,
        audit=container._audit,  # type: ignore[attr-defined]
    )
    print(f"Done. {created} profile rows created.")


if __name__ == "__main__":
    import asyncio

    try:
        asyncio.run(_main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=__import__("sys").stderr)
        raise SystemExit(1) from exc
