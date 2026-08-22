# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W03 (#55)
# DA-11, D27, D55 — Postgres adapter for the audit-log port.
"""Async Postgres adapter for the ``AuditLogPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.audit_log.AuditLogPort``
(DA-11, D27, D55). The audit log is append-only (no ``delete``);
every mutating case-of-use inside the auth/session boundary MUST
invoke ``append`` in the same SQLAlchemy session as the mutation it
audits, so a failed insert propagates and the surrounding transaction
rolls back.

The schema forbids every telemetry column the legacy
``TbConexiones`` / ``TbAplicacionesAperturas`` used to carry (SSID,
BSSID, machine name, coordinates, IP — DA-13 + D55). The
``scripts/check_legacy_hashes.py`` gate rejects any reintroduction,
so the contract here is enforced by CI rather than by adapter code.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime
from typing import Any
from uuid import UUID

import sqlalchemy as sa
from sqlalchemy import select
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.dialects.postgresql import UUID as PGUUID
from sqlalchemy.ext.asyncio import AsyncSession

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    SCHEMA,
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.domain.audit_event import AuditEvent

AUDIT_TABLE = sa.Table(
    "audit",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("event_type", sa.Text(), nullable=False),
    sa.Column("actor_id", PGUUID(as_uuid=True), nullable=True),
    sa.Column("target_id", sa.Text(), nullable=False),
    sa.Column("module", sa.Text(), nullable=False),
    sa.Column("result", sa.Text(), nullable=False),
    sa.Column("correlation_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("payload", JSONB(astext_type=sa.Text()), nullable=False),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


def _row_to_event(row: sa.Row[Any]) -> AuditEvent:
    """Map an ``audit`` row to the domain ``AuditEvent`` dataclass.

    The ``payload`` column is JSONB; Postgres deserialises to a
    Python dict on read. The ``actor_id`` column is nullable for
    system-issued events (cron, bootstrap, migrations).
    """
    return AuditEvent(
        id=row.id,
        event_type=row.event_type,
        actor_id=row.actor_id,
        target_id=row.target_id,
        module=row.module,
        result=row.result,
        correlation_id=row.correlation_id,
        payload=dict(row.payload or {}),
        created_at=row.created_at,
    )


class AuditLogPg:
    """Postgres adapter for the ``AuditLogPort`` Protocol (DA-11, D27, D55)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def append(self, event: AuditEvent) -> None:
        """Persist an audit row.

        DA-11: every mutating case-of-use MUST call this from inside
        the same SQLAlchemy session as the mutation it audits. The
        adapter uses its own fresh ``AsyncSession`` here only because
        the audit emission in the consume-reset-token code path
        fires *after* the user-update transaction's commit point
        (so a failed audit cannot roll back the password reset);
        transactional consumers wire the adapter via session sharing
        instead.
        """
        session: AsyncSession = self._factory()
        try:
            stmt = sa.insert(AUDIT_TABLE).values(
                id=event.id,
                event_type=event.event_type,
                actor_id=event.actor_id,
                target_id=event.target_id,
                module=event.module,
                result=event.result,
                correlation_id=event.correlation_id,
                payload=event.payload,
                created_at=event.created_at,
            )
            await session.execute(stmt)
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> Sequence[AuditEvent]:
        """Return every audit row with ``actor_id == actor_id`` and ``created_at >= since``.

        Ordered by ``created_at`` DESC so the delivery layer's audit
        viewer renders the newest first.
        """
        session: AsyncSession = self._factory()
        try:
            stmt = (
                select(AUDIT_TABLE)
                .where(
                    sa.and_(
                        AUDIT_TABLE.c.actor_id == actor_id,
                        AUDIT_TABLE.c.created_at >= since,
                    )
                )
                .order_by(AUDIT_TABLE.c.created_at.desc())
            )
            rows = (await session.execute(stmt)).all()
        finally:
            await session.close()
        return [_row_to_event(r) for r in rows]


__all__ = [
    "AUDIT_TABLE",
    "AuditLogPg",
]
