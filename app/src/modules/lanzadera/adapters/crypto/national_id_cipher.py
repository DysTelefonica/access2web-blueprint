"""AES-GCM implementation of national-id encryption (D88-CA-S4).

Encrypts ``users.dni_encrypted`` with a key fetched from
:class:`SecretManagerPort` (CA-S2). The key never appears in logs, CLI
arguments or exception messages (CA-S4): callers receive opaque bytes
back and pass them to the database adapter.

Layout: ``nonce (12 bytes) || ciphertext || tag (16 bytes)``. The nonce
is generated fresh per call with ``os.urandom``; the secret key is
re-derived per call from the injected secret manager (no in-process
caching here — caching lives in the composition root if ever required).
"""

from __future__ import annotations

import base64
import binascii
import os

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManagerPort

_NONCE_BYTES = 12
_KEY_BYTES = 32  # AES-256


class NationalIdCipher:
    """Encrypt/decrypt national IDs with AES-256-GCM."""

    def __init__(self, secrets: SecretManagerPort, *, key_name: str = "users.dni.cipher") -> None:
        # The key is fetched lazily and held in memory only for the duration
        # of the call. We do not cache it across calls: if the operator
        # rotates the secret, every subsequent encrypt/decrypt sees the new
        # key without restart.
        self._secrets = secrets
        self._key_name = key_name

    def encrypt(self, plain: str) -> bytes:
        """Encrypt ``plain`` and return ``nonce || ciphertext || tag``."""
        key = self._load_key()
        nonce = os.urandom(_NONCE_BYTES)
        body = AESGCM(key).encrypt(nonce, plain.encode("utf-8"), None)
        return nonce + body

    def decrypt(self, blob: bytes) -> str:
        """Decrypt ``nonce || ciphertext || tag`` back to a UTF-8 string."""
        if len(blob) < _NONCE_BYTES:
            raise ValueError("ciphertext too short; refusing to decrypt")
        key = self._load_key()
        nonce = blob[:_NONCE_BYTES]
        body = blob[_NONCE_BYTES:]
        try:
            plain = AESGCM(key).decrypt(nonce, body, None)
        except InvalidTag:
            # CA-S4: do not include the blob, the key, or any tag value in
            # the message — only the structural fact that auth failed.
            raise ValueError("national-id decryption failed") from None
        return plain.decode("utf-8")

    def _load_key(self) -> bytes:
        """Fetch and decode the secret key. Internal — never returned to callers.

        The key is delivered as a URL-safe base64 string (Fernet-style per
        CA-S2 / D25). We require the encoding to be exactly 32 bytes after
        decoding — AES-256-GCM. Non-Fernet inputs (raw bytes, missing
        padding, wrong alphabet) are rejected up front so CA-S4 holds:
        no key material, no hint about the encoding, just a structural
        refusal.
        """
        import re

        raw = self._secrets.get(self._key_name)
        if not re.fullmatch(r"[A-Za-z0-9_\-]+={0,2}", raw):
            # CA-S4: do not echo the raw value back to the caller.
            raise ValueError("PLATFORM_SECRET_KEY is not a Fernet-format string")
        if len(raw) % 4 != 0:
            raise ValueError("PLATFORM_SECRET_KEY is not a Fernet-format string")
        try:
            decoded = base64.urlsafe_b64decode(raw)
        except (binascii.Error, ValueError):
            raise ValueError("PLATFORM_SECRET_KEY is not a Fernet-format string") from None
        if len(decoded) != _KEY_BYTES:
            # Length-only error: do not leak the length we expected.
            raise ValueError("PLATFORM_SECRET_KEY does not decode to a 32-byte AES-256 key")
        return decoded
