"""Contract-conformance test for :class:`NotificationDeliveryPort`.

Pin both implementations against the same Protocol:
- the in-memory ``FakeNotificationDelivery`` (used by the rest of
  the test suite), and
- the Postgres ``MailQueueTableAdapter`` wired to the
  :func:`async_session_factory` seam (DA-1, W06 #47).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime

from app.src.modules.lanzadera.adapters.persistence import (
    MailQueueTableAdapter,
    async_session_factory,
)


@dataclass
class FakeNotificationDelivery:
    """In-memory notification sink -- captures every ``send`` for inspection."""

    sent: list[tuple[str, str, str, datetime]] = field(default_factory=list)

    async def send(self, to: str, subject: str, body: str) -> None:
        self.sent.append((to, subject, body, datetime.now(UTC)))


def test_mail_queue_pg_satisfies_protocol() -> None:
    """Static structural check on the Postgres adapter's method set."""
    expected = {"send"}
    actual = set(dir(MailQueueTableAdapter))
    # The constructor takes a factory seam; we don't instantiate here
    # because the engine is created lazily.
    missing = expected - actual
    assert not missing, f"MailQueueTableAdapter is missing Protocol methods: {missing}"


async def test_fake_send_records_to_subject_body() -> None:
    fake = FakeNotificationDelivery()
    await fake.send("alice@enterprise.test", "Reset password", "https://example/reset/abc")
    await fake.send("bob@enterprise.test", "Welcome", "hi")
    assert len(fake.sent) == 2
    first = fake.sent[0]
    assert first[0] == "alice@enterprise.test"
    assert first[1] == "Reset password"
    assert first[2] == "https://example/reset/abc"
    # The 4th tuple element is the timestamp -- we just assert it is a datetime.
    assert isinstance(first[3], datetime)


async def test_mail_queue_pg_construction() -> None:
    """``MailQueueTableAdapter`` takes the factory seam and is constructable (DA-1).

    We do not connect to Postgres here; an unreachable URL is fine
    because the engine is created lazily.
    """
    _, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    adapter = MailQueueTableAdapter(factory)
    assert adapter is not None
