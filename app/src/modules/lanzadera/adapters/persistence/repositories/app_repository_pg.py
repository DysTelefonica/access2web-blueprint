# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W01 (#44)
# DA-1, D52, D85 — Postgres adapter for the app port.
"""Async Postgres adapter for the ``AppRepositoryPort``.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.app_repository.AppRepositoryPort``
(DA-7, D52). The hot path ``list_active`` is the
``GET /admin/apps`` query; DA-8 caches it with a 5-minute TTL at a
higher layer (this adapter does NOT cache — the caching decorator
sits between the delivery route and this object).

The ``list_visible_to`` query joins ``apps`` with the
``user_app_assignments`` table; the join's column order is
important for the existing RLS policies (see migration 0001).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

import sqlalchemy as sa
from sqlalchemy import select
from sqlalchemy.dialects.postgresql import UUID

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    SCHEMA,
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.domain.app import (
    App,
    AppRegistrationStatus,
    AppTopology,
)

APPS_TABLE = sa.Table(
    "apps",
    sa.MetaData(),
    sa.Column("id", sa.Integer, sa.Identity(always=False), primary_key=True),
    sa.Column("name", sa.Text(), nullable=False),
    sa.Column("short_code", sa.Text(), nullable=False),
    sa.Column("deployment_topology", sa.Enum(AppTopology, name="app_topology")),
    sa.Column("requires_office_presence", sa.Boolean, nullable=False, default=False),
    sa.Column(
        "registration_status",
        sa.Enum(AppRegistrationStatus, name="app_registration_status"),
    ),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)

USER_APP_ASSIGNMENTS_TABLE = sa.Table(
    "user_app_assignments",
    sa.MetaData(),
    sa.Column("id", UUID(as_uuid=True), primary_key=True),
    sa.Column("user_id", UUID(as_uuid=True), nullable=False),
    sa.Column("app_id", sa.Integer, nullable=False),
    sa.Column("profile_id", UUID(as_uuid=True), nullable=False),
    schema=SCHEMA,
)


def _row_to_app(row: sa.Row[Any]) -> App:
    """Map an ``apps`` row to the domain ``App`` dataclass."""
    return App(
        id=row.id,
        name=row.name,
        short_code=row.short_code,
        deployment_topology=AppTopology(row.deployment_topology),
        requires_office_presence=row.requires_office_presence,
        registration_status=AppRegistrationStatus(row.registration_status),
        created_at=row.created_at,
        updated_at=row.updated_at,
    )


class AppRepositoryPg:
    """Postgres adapter for the ``AppRepositoryPort`` Protocol."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def get_by_id(self, app_id: int) -> App | None:
        """Return the app with ``id == app_id`` or ``None``."""
        async with self._factory.read_only_session() as session:
            stmt = select(APPS_TABLE).where(APPS_TABLE.c.id == app_id)
            row = (await session.execute(stmt)).first()
        return _row_to_app(row) if row is not None else None

    async def list_active(self) -> Sequence[App]:
        """Return every app with ``registration_status='active'``.

        Ordered by ``short_code`` so the delivery layer's `select`
        dropdown is deterministic across renders.
        """
        async with self._factory.read_only_session() as session:
            stmt = (
                select(APPS_TABLE)
                .where(APPS_TABLE.c.registration_status == AppRegistrationStatus.ACTIVE.value)
                .order_by(APPS_TABLE.c.short_code)
            )
            rows = (await session.execute(stmt)).all()
        return [_row_to_app(r) for r in rows]

    async def list_visible_to(self, user_id: object) -> Sequence[App]:
        """Return every app that ``user_id`` is allowed to see.

        Includes both ``active`` rows the user can reach and any
        ``pending`` row they have an explicit ``user_app_assignments``
        grant for. The join's column order matches the indexes in
        migration 0001 so the planner can pick ``ix_assignments_user``
        directly.
        """
        async with self._factory.read_only_session() as session:
            stmt = (
                select(APPS_TABLE)
                .outerjoin(
                    USER_APP_ASSIGNMENTS_TABLE,
                    USER_APP_ASSIGNMENTS_TABLE.c.app_id == APPS_TABLE.c.id,
                )
                .where(
                    sa.or_(
                        APPS_TABLE.c.registration_status == AppRegistrationStatus.ACTIVE.value,
                        sa.and_(
                            USER_APP_ASSIGNMENTS_TABLE.c.user_id == user_id,
                            APPS_TABLE.c.registration_status == AppRegistrationStatus.PENDING.value,
                        ),
                    )
                )
                .order_by(APPS_TABLE.c.short_code)
                .distinct()
            )
            rows = (await session.execute(stmt)).all()
        return [_row_to_app(r) for r in rows]


__all__ = ["AppRepositoryPg", "APPS_TABLE"]
