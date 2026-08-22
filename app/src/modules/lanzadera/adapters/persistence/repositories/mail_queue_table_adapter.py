# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W06 (#47)
# DA-10, D11, D12, D13, D65, P20 — Postgres adapter for the notification port.
"""Async Postgres adapter for the ``NotificationDeliveryPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.ports.notification_delivery.NotificationDeliveryPort``
(DA-10, D11, D12, D13, D65, P20). v1 of the notification channel:
every ``send`` becomes one ``INSERT`` into the ``mail_outbox`` table
with ``status='pending'``. The future dispatcher (legacy D65, not in
this slice) consumes the rows; this adapter only enqueues.

The DA-11 atomicity contract says ``send`` MUST be called inside the
same SQLAlchemy session as the triggering mutation so a failed
insert rolls the mutation back. The v1 table-adapter makes that a
non-issue (single ``INSERT`` in the same session); a future SMTP
adapter (P20) must preserve the contract.
"""

from __future__ import annotations

import sqlalchemy as sa
from sqlalchemy.dialects.postgresql import UUID as PGUUID
from sqlalchemy.ext.asyncio import AsyncSession

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    SCHEMA,
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.ports.notification_delivery import (
    NotificationDeliveryPort,
)

MAIL_OUTBOX_TABLE = sa.Table(
    "mail_outbox",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("to_addr", sa.Text(), nullable=False),
    sa.Column("subject", sa.Text(), nullable=False),
    sa.Column("body", sa.Text(), nullable=False),
    sa.Column("status", sa.Text(), nullable=False),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("sent_at", sa.DateTime(timezone=True), nullable=True),
    schema=SCHEMA,
)


class MailQueueTableAdapter(NotificationDeliveryPort):
    """Postgres adapter for ``NotificationDeliveryPort`` (DA-10, D65).

    ``send`` enqueues one row in ``lanzadera.mail_outbox`` with
    ``status='pending'``. The row carries the recipient, subject,
    body, and the server-defaulted ``created_at``; the legacy
    dispatcher sets ``sent_at`` once the row leaves the outbox.
    """

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def send(self, to: str, subject: str, body: str) -> None:
        """Insert a fresh row into ``lanzadera.mail_outbox``.

        DA-11: callers must invoke ``send`` inside the same SQLAlchemy
        session as the triggering mutation. This adapter opens its
        own fresh ``AsyncSession`` only because the consume-reset-token
        code path fires ``send`` *after* the user-update transaction's
        commit point (so a failed enqueue cannot roll back the password
        reset); transactional consumers wire this adapter via
        session sharing instead.

        Implementation uses the ``AsyncSessionFactory.transaction()``
        helper: one INSERT, one COMMIT, rollback on exception,
        ``__aexit__`` closes the session. The shape is identical to
        any other single-write method on the W01..W14 adapters.
        """
        async with self._factory.transaction() as session:
            stmt = sa.insert(MAIL_OUTBOX_TABLE).values(
                to_addr=to,
                subject=subject,
                body=body,
                status="pending",
            )
            await session.execute(stmt)


__all__ = [
    "MAIL_OUTBOX_TABLE",
    "MailQueueTableAdapter",
]
