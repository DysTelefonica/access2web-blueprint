"""Protocol-conformance test for the ``PasswordHasher`` port (DA-2, D88).

The contract is the ``PasswordHasher`` Protocol declared in
``app.src.modules.lanzadera.domain.ports``. This test pins the concrete
adapter (``CredentialHasherArgon2id``) against the Protocol interface
so the wiring is explicit rather than implicit.

W07 note: ``CredentialHasherArgon2id`` carries two ``async def``
methods (``hash``, ``verify``). The Argon2id KDF itself is CPU-bound
but delegated via ``asyncio.to_thread`` so the event loop stays
responsive. The 100% coverage target (DA-2, CRITICAL_HELPER) is
pinned by ``tests/lanzadera/adapters/crypto/test_credential_hasher_argon2id.py``.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.domain.ports import PasswordHasher


async def test_argon2id_adapter_satisfies_protocol() -> None:
    """The adapter satisfies the ``PasswordHasher`` Protocol structurally."""
    # Static: mypy verifies the structural conformance at module
    # import time. Runtime: an instance check would require the
    # Protocol to be decorated with ``@runtime_checkable``; the
    # typing.Protocol contract is what the composition root uses.
    adapter: PasswordHasher = CredentialHasherArgon2id()
    plain = "protocol-conformance-test"
    hashed = await adapter.hash(plain)
    assert await adapter.verify(plain, hashed) is True
