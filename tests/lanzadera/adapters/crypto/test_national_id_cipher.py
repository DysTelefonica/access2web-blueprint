"""CRITICAL_HELPER tests for :class:`NationalIdCipher` (CA-S4).

Coverage target: 100 % for the cipher adapter. The fake
:class:`FakeSecretManager` lives here (not in a shared ``_fakes.py``)
because the cipher is the only consumer yet — promote it upward when a
second adapter needs an in-memory secret manager.
"""

from __future__ import annotations

import base64
import os

import pytest

from app.src.modules.lanzadera.adapters.crypto.national_id_cipher import (
    NationalIdCipher,
)
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManagerPort


class FakeSecretManager:
    """In-memory :class:`SecretManagerPort` for tests."""

    def __init__(self, **secrets: str) -> None:
        self._store = dict(secrets)

    def get(self, key: str) -> str:
        if key not in self._store:
            raise KeyError(key)
        return self._store[key]

    def set(self, key: str, value: str) -> None:
        self._store[key] = value


@pytest.fixture
def aes_key_bytes() -> bytes:
    return os.urandom(32)


@pytest.fixture
def aes_key_b64(aes_key_bytes: bytes) -> str:
    return base64.urlsafe_b64encode(aes_key_bytes).decode("ascii")


@pytest.fixture
def cipher(aes_key_b64: str) -> NationalIdCipher:
    return NationalIdCipher(
        secrets=FakeSecretManager(**{"users.dni.cipher": aes_key_b64}),
    )


class TestRoundtrip:
    """Encrypt → decrypt returns the original plaintext."""

    def test_roundtrip_simple(self, cipher: NationalIdCipher) -> None:
        plain = "12345678Z"
        blob = cipher.encrypt(plain)
        assert cipher.decrypt(blob) == plain

    @pytest.mark.parametrize("plain", ["", "A", "12345678Z", "X-1234567-Y"])
    def test_roundtrip_various_ids(self, cipher: NationalIdCipher, plain: str) -> None:
        blob = cipher.encrypt(plain)
        assert cipher.decrypt(blob) == plain

    def test_unicode_roundtrip(self, cipher: NationalIdCipher) -> None:
        plain = "ñ-12345-😀"
        blob = cipher.encrypt(plain)
        assert cipher.decrypt(blob) == plain


class TestNonceRandomness:
    """Two encrypts of the same plaintext produce different ciphertexts (nonce random)."""

    def test_same_plaintext_yields_different_ciphertexts(self, cipher: NationalIdCipher) -> None:
        plain = "12345678Z"
        first = cipher.encrypt(plain)
        second = cipher.encrypt(plain)
        assert first != second
        # But both decrypt to the same plaintext.
        assert cipher.decrypt(first) == plain
        assert cipher.decrypt(second) == plain


class TestTamperResistance:
    """AES-GCM authenticates — any bit flip must fail decryption."""

    def test_flipped_ciphertext_byte_raises(self, cipher: NationalIdCipher) -> None:
        blob = cipher.encrypt("12345678Z")
        tampered = bytearray(blob)
        tampered[-1] ^= 0x01  # flip the last bit of the tag
        with pytest.raises(ValueError, match="decryption failed"):
            cipher.decrypt(bytes(tampered))

    def test_flipped_nonce_byte_raises(self, cipher: NationalIdCipher) -> None:
        blob = cipher.encrypt("12345678Z")
        tampered = bytearray(blob)
        tampered[0] ^= 0x01
        with pytest.raises(ValueError, match="decryption failed"):
            cipher.decrypt(bytes(tampered))

    def test_truncated_blob_raises(self, cipher: NationalIdCipher) -> None:
        with pytest.raises(ValueError, match="too short"):
            cipher.decrypt(b"short")


class TestSecretManagerContract:
    """The cipher integrates with :class:`SecretManagerPort` correctly."""

    def test_missing_secret_key_raises_keyerror(self) -> None:
        secrets = FakeSecretManager()  # no secrets configured
        cipher = NationalIdCipher(secrets=secrets)
        with pytest.raises(KeyError):
            cipher.encrypt("any-plain")

    def test_non_fernet_key_rejected(self) -> None:
        # CA-S4: ambiguous encoding is a footgun — we accept only URL-safe
        # base64 (Fernet-style per CA-S2). `!` is not in the base64 alphabet,
        # so the regex rejects the input outright.
        secrets = FakeSecretManager(**{"users.dni.cipher": "!" * 32})
        cipher = NationalIdCipher(secrets=secrets)
        with pytest.raises(ValueError, match="Fernet"):
            cipher.encrypt("plain")

    def test_wrong_decoded_length_raises_value_error(self) -> None:
        # Valid base64 of 24 bytes input → 18 bytes decoded → wrong for AES-256.
        # (32 base64 chars decode to 24 bytes; we want to detect that case.)
        short_key = base64.urlsafe_b64encode(b"x" * 16).decode("ascii")  # 24 chars → 12 bytes
        secrets = FakeSecretManager(**{"users.dni.cipher": short_key})
        cipher = NationalIdCipher(secrets=secrets)
        with pytest.raises(ValueError, match="32-byte"):
            cipher.encrypt("plain")

    def test_works_with_real_port_protocol(self, aes_key_b64: str) -> None:
        # Sanity: the cipher only needs a duck-typed ``get``. The Protocol
        # is structural — assert via static type and via runtime use.
        secrets: SecretManagerPort = FakeSecretManager(**{"users.dni.cipher": aes_key_b64})
        cipher = NationalIdCipher(secrets=secrets)
        assert cipher.decrypt(cipher.encrypt("X")) == "X"


class TestKeyRotation:
    """A new cipher instance with a new key cannot decrypt old blobs."""

    def test_rotated_key_cannot_decrypt_previous_ciphertext(self, aes_key_b64: str) -> None:
        old = NationalIdCipher(secrets=FakeSecretManager(**{"users.dni.cipher": aes_key_b64}))
        blob = old.encrypt("12345678Z")

        new_key = base64.urlsafe_b64encode(os.urandom(32)).decode("ascii")
        new = NationalIdCipher(secrets=FakeSecretManager(**{"users.dni.cipher": new_key}))
        with pytest.raises(ValueError, match="decryption failed"):
            new.decrypt(blob)
