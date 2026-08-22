"""NotificationDeliveryPort.

D11, D12, D13, D65. Driven port for emitting one outbound message.
The MVP adapter is :class:`MailQueueTableAdapter` (PR #47 ships
the in-memory fake; the Postgres adapter lands with M02's mail-outbox
work). The reset flow's ``issue_reset_token`` and the admin
``disable_user`` use this port to send the corresponding email.
"""

from __future__ import annotations

from typing import Protocol


class NotificationDeliveryPort(Protocol):
    """Driven port for one outbound message."""

    async def send(self, to: str, subject: str, body: str) -> None:
        """Persist or send one notification.

        The exact channel is an adapter detail (``MailQueueTableAdapter``
        for v1). Callers MUST invoke ``send`` inside the same SQLAlchemy
        session as the triggering mutation so a failed send rolls the
        mutation back (DA-11).
        """
        ...
