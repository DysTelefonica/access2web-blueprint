"""Argon2id implementation of :class:`CredentialHasherPort` (D88, DA-2).

Profile pinned by the auth-core/spec.md §Argon2id-only credential storage
requirement: ``argon2-cffi==25.1.0`` with ``RFC_9106_LOW_MEMORY``
(Argon2id, m=65536, t=3, p=4). The Argon2id parameters are embedded in
the PHC string — verification re-reads them and the verifier never has
to know the parameters out-of-band.

CRITICAL_HELPER (DA-2, QC-5): covered at 100 % by
``tests/lanzadera/adapters/crypto/test_credential_hasher_argon2id.py``.
Two-method surface, no I/O, no ports — pure function of the Argon2id KDF.
"""

from __future__ import annotations

from argon2 import PasswordHasher
from argon2.exceptions import (
    InvalidHashError,
    VerificationError,
    VerifyMismatchError,
)


class CredentialHasherArgon2id:
    """Argon2id PHC strings, profile ``RFC_9106_LOW_MEMORY``."""

    def __init__(self) -> None:
        # PasswordHasher reads RFC_9106_LOW_MEMORY from argon2-cffi==25.1.0:
        # Argon2id, memory_cost=65536 (64 MiB), time_cost=3, parallelism=4.
        # The profile is library-pinned via app/pyproject.toml — changing it
        # requires a release of argon2-cffi, not a code edit.
        self._hasher = PasswordHasher()

    def hash_password(self, plain: str) -> str:
        """Return the Argon2id PHC string for ``plain``."""
        return self._hasher.hash(plain)

    def verify_password(self, plain: str, hashed: str) -> bool:
        """Return ``True`` iff ``plain`` matches ``hashed``.

        Rejects ``hashed in (None, "")`` before touching Argon2id so the
        seeded-user case (``password_hash = NULL``) cannot trigger a
        KDF run against missing material (auth-core/spec.md §Seeded
        users start with reset required).
        """
        if not hashed:
            raise ValueError("hashed credential is empty; refusing to verify")
        try:
            return self._hasher.verify(hashed, plain)
        except (VerifyMismatchError, VerificationError, InvalidHashError):
            return False
