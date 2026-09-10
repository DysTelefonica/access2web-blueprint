# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W01 (#44)
# DA-12, D22 — Postgres adapter for the profile port.
"""Async Postgres adapter for the ``ProfileRepositoryPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.profile_repository.ProfileRepositoryPort``
(DA-12, D22, D45). The ``capabilities`` column lives as JSONB in the
``lanzadera.profiles`` table; the adapter marshals the Python
``dict[str, str | int | bool]`` into and out of the JSONB wire format.

Profiles are scoped under an ``app_id``. The ``list_for_app`` and
``get_by_code`` methods preserve the constraint; ``create`` enforces
DA-12 by serialising the capabilities map atomically per call.
"""

from __future__ import annotations

from uuid import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    Any,
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.profile_table import (
    PROFILES_TABLE,
)  # noqa: F401  # re-exported for back-compat
from app.src.modules.lanzadera.domain.profile import Profile


def _row_to_profile(row: sa.Row[Any]) -> Profile:
    """Map a ``profiles`` row to the domain ``Profile`` dataclass.

    The ``capabilities`` JSONB column deserialises to a Python dict
    on read; Postgres handles the JSONB ↔ Python transition
    transparently for ``str | int | bool`` leaves. The ``id`` column
    is server-defaulted (``gen_random_uuid()``); for callers that
    construct a ``Profile`` from a row that has no id (the tests'
    in-memory variant), the dataclass enforces the field is present.
    """
    return Profile(
        id=row.id,
        app_id=row.app_id,
        code=row.code,
        name=row.name,
        capabilities=dict(row.capabilities or {}),
        active=row.active,
        created_at=row.created_at,
        updated_at=row.updated_at,
    )


class ProfileRepositoryPg:
    """Postgres adapter for the ``ProfileRepositoryPort`` Protocol (DA-12)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def list_for_app(self, app_id: int) -> Sequence[Profile]:
        """Return every profile of ``app_id`` ordered by ``code``.

        DA-12: the canonical per-app profile set is OPEN (gap G-2);
        the delivery layer filters on ``active`` further down.
        """
        async with self._factory.read_only_session() as session:
            stmt = (
                select(PROFILES_TABLE)
                .where(PROFILES_TABLE.c.app_id == app_id)
                .order_by(PROFILES_TABLE.c.code)
            )
            rows = (await session.execute(stmt)).all()
        return [_row_to_profile(r) for r in rows]

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:
        """Return the profile identified by ``(app_id, code)`` or ``None``."""
        async with self._factory.read_only_session() as session:
            stmt = select(PROFILES_TABLE).where(
                sa.and_(
                    PROFILES_TABLE.c.app_id == app_id,
                    PROFILES_TABLE.c.code == code,
                )
            )
            row = (await session.execute(stmt)).first()
        return _row_to_profile(row) if row is not None else None

    async def get_by_id(self, profile_id: UUID) -> Profile | None:
        """Return the profile identified by ``profile_id`` or ``None``."""
        async with self._factory.read_only_session() as session:
            stmt = select(PROFILES_TABLE).where(
                PROFILES_TABLE.c.id == profile_id,
            )
            row = (await session.execute(stmt)).first()
        return _row_to_profile(row) if row is not None else None

    async def create(self, profile: Profile) -> None:
        """Persist a new profile row.

        Uses the JSONB ``capabilities`` column directly; the
        canonical per-app shape lands in migration 0003 (PR 3b).
        """
        async with self._factory.transaction() as session:
            stmt = sa.insert(PROFILES_TABLE).values(
                app_id=profile.app_id,
                code=profile.code,
                name=profile.name,
                capabilities=profile.capabilities,
                active=profile.active,
            )
            await session.execute(stmt)

    async def set_active(self, app_id: int, code: str, *, active: bool) -> None:
        """Toggle ``active`` for the profile ``(app_id, code)``.

        The capability set is unaffected. ``updated_at`` is touched by
        the Postgres trigger declared in migration 0001.
        """
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(PROFILES_TABLE)
                .where(
                    sa.and_(
                        PROFILES_TABLE.c.app_id == app_id,
                        PROFILES_TABLE.c.code == code,
                    )
                )
                .values(active=active)
            )
            await session.execute(stmt)


__all__ = ["ProfileRepositoryPg", "PROFILES_TABLE"]
