"""NotificationDeliveryPort — D11, D12, D13, D65, P20.

Driven port for emitting a notification to one recipient. v1 persists
the row in ``mail_outbox`` via ``MailQueueTableAdapter``; v2 (open —
P20) will swap the adapter for an SMTP direct path without changing
this contract. The reset flow's ``issue_reset_token`` emits a reset
URL through this port so the application layer never touches the
channel directly.

Delivery semantics: ``send`` MUST be called inside the same SQLAlchemy
session as the mutation that triggered it so a failed insert rolls the
mutation back (DA-11). The v1 table-adapter makes that a non-issue
(it is one ``INSERT`` in the same session), but a future SMTP adapter
must preserve the atomicity contract.
"""

from __future__ import annotations

from typing import Protocol


class NotificationDeliveryPort(Protocol):
    """Driven port for one outbound message."""

    async def send(self, to: str, subject: str, body: str) -> None:
        """Persist or send one notification.

        The exact channel is an adapter detail (``MailQueueTableAdapter``
        for v1). Callers MUST invoke ``send`` inside the same SQLAlchemy
        session as the triggering mutation so a failure propagates and
        the surrounding transaction rolls back (DA-11).
        """
        ...
