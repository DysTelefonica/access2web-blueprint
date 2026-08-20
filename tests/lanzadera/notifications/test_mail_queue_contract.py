"""Contract-conformance test for :class:`NotificationDeliveryPort`.

The ``MailQueueTableAdapter`` body is a skeleton (``NotImplementedError``)
until the runtime harness wires ``mail_outbox`` into the testable
boundary. The in-memory fake is the contract the application layer
relies on today.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from datetime import UTC, datetime

import pytest

from app.src.modules.lanzadera.adapters.notification.mail_queue_table_adapter import (
    MailQueueTableAdapter,
)
from app.src.modules.lanzadera.ports.notification_delivery import (
    NotificationDeliveryPort,
)


@dataclass
class FakeNotificationDelivery:
    """In-memory notification sink — captures every ``send`` for inspection."""

    sent: list[tuple[str, str, str, datetime]] = field(default_factory=list)

    async def send(self, to: str, subject: str, body: str) -> None:
        self.sent.append((to, subject, body, datetime.now(UTC)))


def test_mail_queue_pg_satisfies_protocol() -> None:
    adapter: NotificationDeliveryPort = MailQueueTableAdapter()
    assert hasattr(adapter, "send")


def test_fake_satisfies_protocol() -> None:
    """Static + runtime: the fake implements the same Protocol as the adapter."""
    fake: NotificationDeliveryPort = FakeNotificationDelivery()
    assert hasattr(fake, "send")


def test_fake_send_records_to_subject_body() -> None:
    fake = FakeNotificationDelivery()

    async def _go() -> None:
        await fake.send("[email protected]", "Reset password", "https://example/reset/abc")
        await fake.send("[email protected]", "Welcome", "hi")
        assert len(fake.sent) == 2
        first = fake.sent[0]
        assert first[0] == "[email protected]"
        assert first[1] == "Reset password"
        assert first[2] == "https://example/reset/abc"
        # The 4th tuple element is the timestamp — we just assert it is a datetime.
        assert isinstance(first[3], datetime)

    asyncio.run(_go())


def test_fake_send_does_not_call_other_methods() -> None:
    """The port contract has exactly one method; the adapter must not invent others."""
    adapter: NotificationDeliveryPort = MailQueueTableAdapter()
    public = [m for m in dir(adapter) if not m.startswith("_")]
    assert public == ["send"], f"unexpected public surface: {public}"


def test_pg_skeleton_raises_not_implemented() -> None:
    """The Pg adapter is a skeleton — calling ``send`` raises.

    This is the contract today: the body awaits the mail_outbox
    testable boundary. The in-memory fake covers the application
    layer's needs in the meantime.
    """
    adapter = MailQueueTableAdapter()

    async def _go() -> None:
        await adapter.send("[email protected]", "x", "y")

    with pytest.raises(NotImplementedError):
        asyncio.run(_go())
