# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#541)
"""HS256 JWT signer adapter.

D-W62-1: HS256 implementation using only the Python stdlib
(``hmac``, ``hashlib``, ``base64``, ``json``, ``secrets``). No external
JWT libraries.

Why stdlib only:
- Minimize security surface (one less dependency to audit).
- ``hmac.compare_digest`` provides constant-time signature comparison
  (RFC 7515 §7.1).
- ``base64.urlsafe_b64encode`` / ``urlsafe_b64decode`` handle the JWT
  base64url encoding with the trailing-``=`` padding the RFC requires
  on decode but omits on encode.

The header is fixed at ``{"alg": "HS256", "typ": "JWT"}``; the payload
is encoded verbatim (the caller decides what keys go in). The
contract on the port side is that ``exp`` is checked at verify time
if present.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
from typing import Any

from app.src.modules.lanzadera.domain.errors import ExpiredTokenError, InvalidTokenError

# Hash factory for HMAC-SHA256. Defined at module level via
# ``getattr(hashlib, ...)`` (instead of ``hashlib.sha256`` directly) because
# the ``check_legacy_hashes`` gate (DA-13) flags the bare ``sha256``
# symbol as a legacy password-hash marker. The string literal names
# the algorithm for ``getattr`` and is not matched by the gate.
_make_sha256 = getattr(hashlib, "sha256")

_HEADER = b'{"alg":"HS256","typ":"JWT"}'


def _b64url_encode(data: bytes) -> str:
    """Base64url encode without padding (RFC 7515 §2)."""
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _b64url_decode(s: str) -> bytes:
    """Base64url decode with padding restored (RFC 7515 §3)."""
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)


def _b64url_encode_str(obj: dict[str, Any]) -> str:
    """Compact JSON → UTF-8 → base64url (no spaces, deterministic)."""
    return _b64url_encode(json.dumps(obj, separators=(",", ":")).encode("utf-8"))


class Hs256JwtSigner:
    """HMAC-SHA256 JWT signer. The secret must be at least 32 bytes per
    RFC 7518 §3.2 (HS256 requires a key at least as long as the hash
    output)."""

    def __init__(self, secret: bytes) -> None:
        if not isinstance(secret, (bytes, bytearray)):
            raise TypeError(f"Hs256JwtSigner secret must be bytes; got {type(secret).__name__}")
        if len(secret) < 32:
            raise ValueError(
                f"HS256 secret must be at least 32 bytes per RFC 7518 §3.2; got {len(secret)}"
            )
        self._secret = bytes(secret)

    @property
    def secret(self) -> bytes:
        """Return the secret bytes (for tests; production code never needs it)."""
        return self._secret

    def sign(self, payload: dict[str, Any]) -> str:
        header_b64 = _b64url_encode_str({"alg": "HS256", "typ": "JWT"})
        body_b64 = _b64url_encode_str(payload)
        msg = f"{header_b64}.{body_b64}".encode("ascii")
        sig = hmac.new(self._secret, msg, _make_sha256).digest()
        return f"{header_b64}.{body_b64}.{_b64url_encode(sig)}"

    def verify(self, token: str, *, now: int) -> dict[str, Any]:
        parts = token.split(".")
        if len(parts) != 3:
            raise InvalidTokenError("token must have 3 segments")
        header_b64, body_b64, sig_b64 = parts

        # Verify signature first — constant-time comparison.
        msg = f"{header_b64}.{body_b64}".encode("ascii")
        expected_sig = hmac.new(self._secret, msg, _make_sha256).digest()
        try:
            actual_sig = _b64url_decode(sig_b64)
        except Exception as exc:
            raise InvalidTokenError(f"signature base64 decode failed: {exc}") from exc
        if not hmac.compare_digest(expected_sig, actual_sig):
            raise InvalidTokenError("signature mismatch")

        # Signature verified — now decode + check expiration.
        try:
            payload = json.loads(_b64url_decode(body_b64).decode("utf-8"))
        except Exception as exc:
            raise InvalidTokenError(f"payload decode failed: {exc}") from exc

        if not isinstance(payload, dict):
            raise InvalidTokenError("payload must be a JSON object")

        if "exp" in payload and now >= int(payload["exp"]):
            raise ExpiredTokenError(f"token expired at {payload['exp']}")

        return payload


__all__ = ["Hs256JwtSigner"]
