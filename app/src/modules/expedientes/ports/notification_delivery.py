"""NotificationDeliveryPort — D-EXP-5, DA-10."""

from __future__ import annotations

from typing import Protocol


class NotificationDeliveryPort(Protocol):
    """Notificaciones al usuario. v1 persiste en mail_outbox; v2 swap SMTP sin tocar dominio."""

    async def send(self, to: str, subject: str, body: str) -> None: ...
