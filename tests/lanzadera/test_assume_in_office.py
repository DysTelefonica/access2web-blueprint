"""Contract-conformance test for :class:`LocationPort` + :class:`AssumeInOfficeAdapter`.

DA-9, H12, D53, D75. The MVP adapter is a stub that always returns
``True``, which disables the office-only gate everywhere in the
application. The real implementation (corporate-IP detection, NAC,
geofencing) lands when the HPS module ships — the first app with
``requires_office_presence`` set in production.

This test pins the stub contract: a future PR that flips the stub to
``False`` (e.g. by introducing a corporate-IP check) is caught by the
test suite, forcing the application layer to update the gate.
"""

from __future__ import annotations

import asyncio
from uuid import uuid4

from app.src.modules.lanzadera.adapters.cross.location import AssumeInOfficeAdapter
from app.src.modules.lanzadera.ports.location import LocationPort


def test_assume_in_office_satisfies_protocol() -> None:
    """Static: the adapter implements the Protocol."""
    adapter: LocationPort = AssumeInOfficeAdapter()
    assert hasattr(adapter, "is_user_in_office")
    assert callable(adapter.is_user_in_office)


def test_assume_in_office_returns_true_for_any_user() -> None:
    """DA-9: the MVP stub returns ``True`` unconditionally.

    H12 says the real check is a future WU; today the office-only
    gate is bypassed for every user.
    """
    adapter = AssumeInOfficeAdapter()

    async def _go() -> None:
        for _ in range(3):
            user_id = uuid4()
            assert await adapter.is_user_in_office(user_id) is True

    asyncio.run(_go())


def test_assume_in_office_ignores_user_id() -> None:
    """The stub MUST return ``True`` even for a ``uuid4()`` we have never seen."""
    adapter = AssumeInOfficeAdapter()
    fresh = uuid4()

    async def _go() -> bool:
        return await adapter.is_user_in_office(fresh)

    assert asyncio.run(_go()) is True


def test_assume_in_office_does_not_track_calls() -> None:
    """The MVP stub does NOT keep a list of which users were asked.

    That would be a side-effect that the application layer does not
    request and that would leak memory in a long-running process.
    A future real adapter (e.g. one that records a 'last seen' timestamp)
    is the right place to add observability.
    """
    adapter = AssumeInOfficeAdapter()
    assert not hasattr(adapter, "calls")
    assert not hasattr(adapter, "users_seen")
