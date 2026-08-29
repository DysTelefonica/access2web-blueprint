"""CredentialHasher port (DA-2, D88).

Methods are async to match the design (DA-1): the Postgres adapter
uses ``argon2-cffi``'s async bindings under the hood, and the
HTTP delivery that calls them already runs inside an event loop.
The in-memory fakes in ``tests/lanzadera/auth/_fakes.py`` carry the
matching ``async def`` shape so the contract is symmetric end to
end.

D89: the hash is the only persisted credential. ``update_password_and_activate``
stores the new hash and flips ``status`` to ``ACTIVE`` in the same
transaction.
"""

from __future__ import annotations

from typing import Protocol


class PasswordHasherPort(Protocol):
    """Argon2id port (DA-2, D88). Real adapter in PR 46."""

    async def hash(self, password: str) -> str: ...
    async def verify(self, password: str, password_hash: str) -> bool: ...


PasswordHasher = PasswordHasherPort


__all__ = ["PasswordHasher", "PasswordHasherPort"]
