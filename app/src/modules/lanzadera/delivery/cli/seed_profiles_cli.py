# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp
"""CLI runner for the ``seed_profiles`` idempotent seed (DA-12, 0003_seed_profiles).

This module lives in the ``delivery`` layer so it may import the Postgres
adapters directly. The ``seed_profiles`` use case itself stays clean in
``application/seed_profiles`` and accepts only port interfaces.
"""

from __future__ import annotations

import asyncio
import os
import sys

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    async_session_factory,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.audit_log_pg import (
    AuditLogPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.profile_repository_pg import (
    ProfileRepositoryPg,
)
from app.src.modules.lanzadera.application.seed_profiles import seed_profiles


async def main() -> None:
    """Build a real session factory and run the seed.

    Requires:
        DATABASE_URL   — Postgres connection string

    Exits with code 0 on success, 1 on error.
    """
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        print(
            "ERROR: DATABASE_URL environment variable is not set.",
            file=sys.stderr,
        )
        raise SystemExit(1)

    print("Building session factory...")
    _engine, session_factory = async_session_factory(database_url)

    profile_repo = ProfileRepositoryPg(session_factory=session_factory)
    audit_log = AuditLogPg(session_factory=session_factory)

    print("Seeding profiles...")
    created = await seed_profiles(
        profiles=profile_repo,
        audit=audit_log,  # type: ignore[arg-type]
    )
    print(f"Done. {created} profile rows created.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as exc:  # pragma: no cover
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
