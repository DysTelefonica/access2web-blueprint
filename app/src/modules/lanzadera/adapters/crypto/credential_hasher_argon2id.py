"""Argon2id implementation of :class:`PasswordHasher` (D88, DA-2).

Profile pinned by the auth-core/spec.md §Argon2id-only credential storage
requirement: ``argon2-cffi==25.1.0`` with ``RFC_9106_LOW_MEMORY``
(Argon2id, m=65536, t=3, p=4). The Argon2id parameters are embedded in
the PHC string -- verification re-reads them and the verifier never has
to know the parameters out-of-band.

CRITICAL_HELPER (DA-2, QC-5): covered at 100 % by
``tests/lanzadera/adapters/crypto/test_credential_hasher_argon2id.py``.
Two-method surface; the methods are ``async def`` to match the rest of
the port surface (DA-1, every port is async end-to-end). Argon2id
itself is CPU-bound; we delegate via ``asyncio.to_thread`` so the event
loop stays responsive under load (DA-1: every port call lands on an
event loop, even when the underlying work is KDF-bound).
"""

from __future__ import annotations

import asyncio

from argon2 import PasswordHasher as _Argon2PasswordHasher
from argon2.exceptions import (
    InvalidHashError,
    VerificationError,
    VerifyMismatchError,
)


class CredentialHasherArgon2id:
    """Argon2id PHC strings, profile ``RFC_9106_LOW_MEMORY``.

    Implements the ``PasswordHasher`` Protocol declared in
    ``app.src.modules.lanzadera.domain.ports``. Three properties:

    - the hash is the canonical Argon2id PHC string with embedded
      parameters ``m=65536, t=3, p=4``;
    - two invocations with the same plaintext return distinct outputs
      (random salt);
    - ``verify`` rejects empty / null stored hashes before touching the
      KDF (auth-core/spec.md §Seeded users start with reset required).
    """

    def __init__(self) -> None:
        # PasswordHasher reads RFC_9106_LOW_MEMORY from argon2-cffi==25.1.0:
        # Argon2id, memory_cost=65536 (64 MiB), time_cost=3, parallelism=4.
        # The profile is library-pinned via app/pyproject.toml -- changing it
        # requires a release of argon2-cffi, not a code edit.
        self._hasher = _Argon2PasswordHasher()

    async def hash(self, password: str) -> str:
        """Return the Argon2id PHC string for ``password``.

        Async: the KDF run is delegated to ``asyncio.to_thread`` so the
        event loop stays responsive during the 64-MiB memory sweep.
        """
        return await asyncio.to_thread(self._hasher.hash, password)

    async def verify(self, password: str, password_hash: str) -> bool:
        """Return ``True`` iff ``password`` matches ``password_hash``.

        Rejects an empty/null stored hash before touching Argon2id so
        the seeded-user case (``password_hash = NULL``) cannot trigger
        work against the KDF. Verifier exceptions resolve to ``False``
        (no information leak about which leg of the check failed).

        Async: the KDF run is delegated to ``asyncio.to_thread`` for the
        same reason as ``hash``.
        """
        if not password_hash:
            raise ValueError("hashed credential is empty; refusing to verify")
        try:
            return await asyncio.to_thread(self._hasher.verify, password_hash, password)
        except (VerifyMismatchError, VerificationError, InvalidHashError):
            return False
