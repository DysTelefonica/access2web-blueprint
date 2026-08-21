"""``AssumeInOfficeAdapter`` — DA-9, H12, D53, D75.

MVP location stub. The adapter returns ``True`` for every user, which
disables the office-only gate everywhere in the application. The
real implementation (corporate-IP detection, NAC, geofencing) lands
when the HPS module ships — the first app with ``requires_office_presence``
set in production.

The stub is intentionally trivial: the application layer's
``is_user_in_office`` check is bypassed for every user today, and
the test that pins this contract (``tests/lanzadera/test_assume_in_office.py``)
makes the bypass explicit so a future PR that flips the stub to
``False`` is caught by the test suite.
"""

from __future__ import annotations

from uuid import UUID

from app.src.modules.lanzadera.ports.location import LocationPort


class AssumeInOfficeAdapter(LocationPort):
    """MVP location stub — always returns ``True``."""

    async def is_user_in_office(self, user_id: UUID) -> bool:
        """Return True unconditionally (DA-9)."""
        # The ``user_id`` parameter is ignored today; it is part of the
        # contract for the future real adapter.
        del user_id
        return True
