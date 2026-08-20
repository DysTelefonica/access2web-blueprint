"""CRITICAL_HELPER port for credential hashing (DA-2, D88).

Argon2id via ``argon2-cffi==25.1.0`` profile ``RFC_9106_LOW_MEMORY``
(m=65536, t=3, p=4). Two-method surface — synchronous on purpose:

* ``hash_password(plain) -> str`` returns the Argon2id PHC string with
  embedded parameters and a random salt.
* ``verify_password(plain, hashed) -> bool`` accepts the matching
  plaintext and rejects any other input (including ``None`` / empty
  hash — the helper refuses to run Argon2id against an empty stored
  hash so a seeded user with ``password_hash = NULL`` cannot trigger
  work against the KDF).

Adapter: :mod:`app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id`.
"""

from __future__ import annotations

from typing import Protocol


class CredentialHasherPort(Protocol):
    """Argon2id via argon2-cffi==25.1.0 profile RFC_9106_LOW_MEMORY (D88)."""

    def hash_password(self, plain: str) -> str:
        """Return the Argon2id PHC string for ``plain``.

        Two invocations with the same plaintext return distinct outputs
        (random salt). The output is the canonical Argon2id PHC string
        with embedded parameters ``m=65536, t=3, p=4``.
        """
        ...

    def verify_password(self, plain: str, hashed: str) -> bool:
        """Return ``True`` iff ``plain`` matches ``hashed``.

        Raises ``ValueError`` when ``hashed`` is empty or ``None``: the
        helper refuses to run Argon2id against a missing/empty stored
        hash (auth-core/spec.md §Seeded users start with reset required).
        """
        ...
