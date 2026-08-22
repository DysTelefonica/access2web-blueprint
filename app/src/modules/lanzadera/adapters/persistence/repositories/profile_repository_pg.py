"""``ProfileRepositoryPg`` — Postgres adapter for :class:`ProfileRepositoryPort`.

Maps rows from ``lanzadera.profiles`` to the :class:`Profile`
dataclass. The read path is hot: the delivery layer calls
``list_for_app(app_id)`` for every permission check; the cache
wraps this with TTL 5 min (DA-8).

``capabilities`` is JSONB (``dict[str, str | int | bool]``). The
canonical capability set per app is OPEN (``##ABIERTO##``,
gap G-2 in the design); the seed migration 0003 (PR 3b) writes
``{}`` until product confirms the catalogue.

Type annotations: see ``user_repository_pg.py`` for the rationale.
The JSONB column deserialises to a Python dict via the ``JSONB``
type's built-in handler.
"""

from __future__ import annotations

from typing import Any

import sqlalchemy as sa
from sqlalchemy import insert, select, update

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)
from app.src.modules.lanzadera.domain.profile import Profile

SCHEMA = "lanzadera"
TABLE = sa.Table(
    "profiles",
    sa.MetaData(),
    sa.Column("id", sa.UUID(), primary_key=True),
    schema=SCHEMA,
)


def _row_to_profile(row: Any) -> Profile:
    return Profile(
        id=row.id,
        app_id=row.app_id,
        code=row.code,
        name=row.name,
        capabilities=dict(row.capabilities) if row.capabilities else {},
        active=bool(row.active),
        created_at=row.created_at,
        updated_at=row.updated_at,
    )


class ProfileRepositoryPg(ProfileRepositoryPort):
    """Postgres adapter for :class:`ProfileRepositoryPort`."""

    def __init__(self, factory: AsyncSessionFactoryPort) -> None:
        self._factory = factory

    async def list_for_app(self, app_id: int) -> list[Profile]:  # type: ignore[misc]
        session = self._factory()
        try:
            stmt = (
                select(TABLE)
                .where(  # type: ignore[arg-type]
                    (TABLE.c.app_id == app_id) & (TABLE.c.active.is_(True))
                )
                .order_by(TABLE.c.code)
            )
            rows = (await session.execute(stmt)).all()
        finally:
            await session.close()
        return [_row_to_profile(r) for r in rows]

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:  # type: ignore[misc]
        session = self._factory()
        try:
            stmt = select(TABLE).where(  # type: ignore[arg-type]
                (TABLE.c.app_id == app_id) & (TABLE.c.code == code)
            )
            row = (await session.execute(stmt)).first()
        finally:
            await session.close()
        return None if row is None else _row_to_profile(row)

    async def create(self, profile: Profile) -> None:  # type: ignore[misc]
        session = self._factory()
        try:
            stmt = insert(TABLE).values(
                id=profile.id,
                app_id=profile.app_id,
                code=profile.code,
                name=profile.name,
                capabilities=profile.capabilities,
                active=profile.active,
            )
            await session.execute(stmt)
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()

    async def set_active(self, profile_id: Any, active: bool) -> None:  # type: ignore[misc]
        session = self._factory()
        try:
            stmt = (
                update(TABLE)
                .where(TABLE.c.id == profile_id)  # type: ignore[arg-type]
                .values(
                    active=active,
                    updated_at=__import__("datetime").datetime.now(__import__("datetime").UTC),
                )
            )
            await session.execute(stmt)
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()
