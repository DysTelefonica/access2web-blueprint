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

from sqlalchemy.dialects.postgresql import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    Any,
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.app_table import (
    APPS_TABLE,
)  # noqa: F401  # re-exported for back-compat
from app.src.modules.lanzadera.domain.app import (
    App,
    AppRegistrationStatus,
    AppTopology,
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

        async def create(
            self,
            name: str,
            short_code: str,
            deployment_topology: AppTopology,
            requires_office_presence: bool,
        ) -> App:
            """Insert a new catalog row and return it with the DB-authoritative id and timestamps.

            W61 (#524): the admin ``POST /admin/apps`` JSON route is the
            first caller. ``id`` is auto-generated by the SERIAL column;
            ``created_at`` / ``updated_at`` are server-defaulted to
            ``now()`` via migration 0001, so the RETURNING round-trip
            reads the authoritative values back. The adapter stamps
            ``registration_status='pending'`` explicitly — keeping it on
            the wire rather than relying on the server default makes the
            intent visible at the call site (DA-7).
            """
            async with self._factory.transaction() as session:
                stmt = (
                    sa.insert(APPS_TABLE)
                    .values(
                        name=name,
                        short_code=short_code,
                        deployment_topology=deployment_topology.value,
                        requires_office_presence=requires_office_presence,
                        registration_status=AppRegistrationStatus.PENDING.value,
                    )
                    .returning(APPS_TABLE)
                )
                row = (await session.execute(stmt)).first()
            if row is None:
                # RETURNING never returns an empty result for a successful
                # INSERT; raising keeps the adapter aligned with the
                # assignment_repository_pg.create() contract.
                raise RuntimeError(f"app create() returned no row for short_code={short_code!r}")
            return _row_to_app(row)

        async def update(
            self,
            app_id: int,
            *,
            name: str | None = None,
            deployment_topology: AppTopology | None = None,
            requires_office_presence: bool | None = None,
        ) -> App:
            """Apply the partial patch to ``app_id`` and return the new row.

            W61 (#524): the admin ``PATCH /admin/apps/{id}`` route is the
            first caller. ``None`` fields are skipped (the column keeps
            its current value); the touched columns are returned via
            ``RETURNING`` so the caller never sees a stale read. The
            adapter always stamps ``updated_at`` so callers cannot forget
            to bump it (the W-TEST pattern the user / assignment
            adapters already pin).
            """
            from datetime import UTC, datetime

            updates: dict[str, object] = {}
            if name is not None:
                updates["name"] = name
            if deployment_topology is not None:
                updates["deployment_topology"] = deployment_topology.value
            if requires_office_presence is not None:
                updates["requires_office_presence"] = requires_office_presence
            updates["updated_at"] = datetime.now(UTC)
            async with self._factory.transaction() as session:
                stmt = (
                    sa.update(APPS_TABLE)
                    .where(APPS_TABLE.c.id == app_id)
                    .values(**updates)
                    .returning(APPS_TABLE)
                )
                row = (await session.execute(stmt)).first()
            if row is None:
                raise RuntimeError(f"app update() returned no row for id={app_id!r}")
            return _row_to_app(row)

        async def disable(self, app_id: int) -> App:
            """Flip ``app_id`` to ``registration_status='retired'`` and return the row.

            W61 (#524): the admin ``DELETE /admin/apps/{id}`` route is
            the first caller. The adapter also stamps ``updated_at`` so
            the lifecycle timestamp stays in step with the transition
            (mirrors the ``update`` behaviour; the ``created_at`` column
            is untouched). ``retired`` rows are filtered out by
            ``list_active`` (the hot read path).
            """
            from datetime import UTC, datetime

            async with self._factory.transaction() as session:
                stmt = (
                    sa.update(APPS_TABLE)
                    .where(APPS_TABLE.c.id == app_id)
                    .values(
                        registration_status=AppRegistrationStatus.RETIRED.value,
                        updated_at=datetime.now(UTC),
                    )
                    .returning(APPS_TABLE)
                )
                row = (await session.execute(stmt)).first()
            if row is None:
                raise RuntimeError(f"app disable() returned no row for id={app_id!r}")
            return _row_to_app(row)


__all__ = ["AppRepositoryPg", "APPS_TABLE"]
