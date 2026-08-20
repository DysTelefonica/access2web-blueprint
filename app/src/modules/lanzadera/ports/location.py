"""LocationPort — DA-9, H12, D53.

The MVP adapter is :class:`AssumeInOfficeAdapter`, which returns
``True`` unconditionally. The platform's ``apps.requires_office_presence``
flag is declared but not evaluated in the MVP (H12 — topología
híbrida is documented as ABRIERTO). When the HPS module lands it
becomes the first app with office-only enforcement; the real
adapter (corporate-IP detection, NAC, etc.) replaces this stub
without changing the contract.
"""

from __future__ import annotations

from typing import Protocol
from uuid import UUID


class LocationPort(Protocol):
    """Driven port for \"is this user in the office right now?\"."""

    async def is_user_in_office(self, user_id: UUID) -> bool:
        """Return True iff the user is currently in the office.

        The MVP contract is permissive: the stub adapter returns
        ``True`` for every user, which means the application layer's
        office-only gate is bypassed. H12 says the real check is a
        future WU; today this port is wired so the gate can compile
        and the application layer can be tested.
        """
        ...
