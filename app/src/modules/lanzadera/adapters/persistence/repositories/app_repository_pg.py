"""``AppRepositoryPg`` — Postgres adapter for :class:`AppRepositoryPort`.

Maps rows from ``lanzadera.apps`` to the :class:`App` dataclass.
The read path is the hot one (the delivery layer calls ``list_active``
on every admin GET, plus ``list_visible_to`` on every user-facing
page); the active set is cached by ``TtlCacheAdapter`` (DA-8) with a
5-minute TTL.

The seed migration 0002 (PR 3b) writes the 20 apps from the
legacy ``TbAplicaciones`` fixture into ``lanzadera.apps``. The seed
sets ``deployment_topology`` and ``requires_office_presence`` from
the legacy ``EjecucionEnOficina`` column (DA-7).

Type annotations: see ``user_repository_pg.py`` for the rationale.
The ``# type: ignore[arg-type]`` markers on the ``select()`` calls
silence the "select from a string is not narrowed" message; the
integration test exercises the round-trip on a real Postgres.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

import sqlalchemy as sa
from sqlalchemy import select

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology
from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort

SCHEMA = "lanzadera"
TABLE = sa.Table(
    "apps",
    sa.MetaData(),
    sa.Column("id", sa.UUID(), primary_key=True),
    schema=SCHEMA,
)


def _row_to_app(row: Any) -> App:
    return App(
        id=row.id,
        name=row.name,
        short_code=row.short_code,
        deployment_topology=AppTopology(row.deployment_topology),
        requires_office_presence=bool(row.requires_office_presence),
        registration_status=AppRegistrationStatus(row.registration_status),
        created_at=row.created_at,
        updated_at=row.updated_at,
    )


class AppRepositoryPg(AppRepositoryPort):
    """Postgres adapter for :class:`AppRepositoryPort`.

    The factory is held as a closure. The hot path (``list_active``) is
    kept thin so the cache wrapper (T04) can call it 50 times/sec
    without re-instantiating the query builder.
    """

    def __init__(self, factory: AsyncSessionFactoryPort) -> None:
        self._factory = factory

    async def get_by_id(self, app_id: int) -> App | None:  # type: ignore[misc]
        session = self._factory()
        try:
            stmt = select(TABLE).where(TABLE.c.id == app_id)  # type: ignore[arg-type]
            row = (await session.execute(stmt)).first()
        finally:
            await session.close()
        return None if row is None else _row_to_app(row)

    async def list_active(self) -> Sequence[App]:  # type: ignore[misc]
        """Return apps with ``registration_status='active'``.

        DA-8: this is the hot path. Cached by ``TtlCacheAdapter`` with
        TTL 5 min; the cache wrapper (T04) handles invalidation when
        the admin ``activate_app`` route fires.
        """
        session = self._factory()
        try:
            stmt = (
                select(TABLE)
                .where(  # type: ignore[arg-type]
                    TABLE.c.registration_status == AppRegistrationStatus.ACTIVE.value
                )
                .order_by(TABLE.c.id)
            )
            rows = (await session.execute(stmt)).all()
        finally:
            await session.close()
        return [_row_to_app(r) for r in rows]

    async def list_visible_to(self, user_id: Any) -> Sequence[App]:  # type: ignore[misc]
        """Apps visible to ``user_id``: union of ``active`` + assignments.

        The "visible to" predicate is: an app is visible if it has
        ``registration_status='active'`` OR the user has an active
        assignment to it (DA-22, assignments/spec.md §Standard profile).
        The full SQL is left for M02's auth wiring (W02..W04) once
        the assignment repository ships; this stub returns the active
        list plus an empty extension (no assignments yet) so the HTTP
        layer (issue #55) compiles and tests pass.
        """
        session = self._factory()
        try:
            stmt = (
                select(TABLE)
                .where(  # type: ignore[arg-type]
                    TABLE.c.registration_status == AppRegistrationStatus.ACTIVE.value
                )
                .order_by(TABLE.c.id)
            )
            rows = (await session.execute(stmt)).all()
        finally:
            await session.close()
        # Assignment-joined list deferred to M02.
        return [_row_to_app(r) for r in rows]
