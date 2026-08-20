"""``MailQueueTableAdapter`` — DA-10, D11, D13, D65, P20.

v1 of the notification channel: every ``send`` becomes one ``INSERT``
into ``mail_outbox`` with ``status='pending'`` and a future dispatcher
(legacy D65, not in this slice) consumes the rows. The adapter is
the only place that knows the ``mail_outbox`` schema; the domain
sees only ``NotificationDeliveryPort``.

**Skeleton** — the body raises ``NotImplementedError`` until the
``mail_outbox`` table is queryable in the runtime harness (migration
``0001_core_schema`` declares it on main, but the testable boundary
that QC-5 measures does not yet include the row-insert path). The
follow-up WU (after the next migration lands) will fill the body
without changing the contract — the in-memory fake already exercises
the shape the application layer cares about.
"""

from __future__ import annotations

from app.src.modules.lanzadera.ports.notification_delivery import (
    NotificationDeliveryPort,
)


class MailQueueTableAdapter(NotificationDeliveryPort):
    """SQLAlchemy Core + asyncpg implementation. Skeleton — see module docstring."""

    async def send(self, to: str, subject: str, body: str) -> None:
        raise NotImplementedError(
            "MailQueueTableAdapter.send awaits the mail_outbox testable boundary"
        )
