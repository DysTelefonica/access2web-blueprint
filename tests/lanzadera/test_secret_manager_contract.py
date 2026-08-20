"""Protocol-conformance test for :class:`SecretManagerPort`."""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.crypto.national_id_cipher import (
    NationalIdCipher,
)
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManagerPort
from tests.lanzadera.adapters.crypto.test_national_id_cipher import FakeSecretManager


def test_fake_secret_manager_satisfies_protocol() -> None:
    # NationalIdCipher consumes SecretManagerPort structurally.
    import base64

    key = base64.urlsafe_b64encode(b"x" * 32).decode("ascii")  # Fernet-format 32-byte AES key
    secrets: SecretManagerPort = FakeSecretManager(**{"users.dni.cipher": key})
    cipher = NationalIdCipher(secrets=secrets)
    blob = cipher.encrypt("protocol-conformance")
    assert cipher.decrypt(blob) == "protocol-conformance"
