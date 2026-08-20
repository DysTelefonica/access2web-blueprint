"""Protocol-conformance test for :class:`CredentialHasherPort`.

The contract is structural — any adapter that satisfies the Protocol is
eligible to be wired into the composition root. This test pins the
concrete adapter (``CredentialHasherArgon2id``) against the Protocol
interface so the wiring is explicit rather than implicit.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.domain.ports.credential_hasher import (
    CredentialHasherPort,
)


def test_argon2id_adapter_satisfies_protocol() -> None:
    # Static: at module import time mypy verifies the structural conformance.
    # Runtime: the assignment would raise TypeError if the Protocol is
    # violated — that is the point of typing.Protocol.
    adapter: CredentialHasherPort = CredentialHasherArgon2id()
    plain = "protocol-conformance-test"
    hashed = adapter.hash_password(plain)
    assert adapter.verify_password(plain, hashed) is True
