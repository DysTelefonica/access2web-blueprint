"""JwtSignerPort — D-W62-1.

Driven port for signing and verifying JWTs. The production adapter
(``Hs256JwtSigner``) uses HMAC-SHA256 from the stdlib — no external
dependency required. The fake (``FakeJwtSigner``) wraps the production
adapter so the test path behaves identically while still recording
call sites for assertion.

D-W62-1: zero external JWT libraries. ``hmac`` + ``hashlib`` +
``base64`` stdlib only. This keeps the security surface minimal and
removes a class of supply-chain risks.

DA-1: the port's ``verify`` takes ``now: int`` (Unix timestamp) so
expiration can be tested deterministically — the production path
passes ``int(time.time())`` and the test path passes a fixed value.
"""

from __future__ import annotations

from typing import Any, Protocol


class JwtSignerPort(Protocol):
    """Sign and verify JWTs with HMAC-SHA256."""

    def sign(self, payload: dict[str, Any]) -> str:
        """Produce a JWT string ``header_b64.payload_b64.sig_b64``.

        The header is fixed to ``{"alg": "HS256", "typ": "JWT"}``. The
        payload is encoded verbatim (caller controls the keys; ``sub``,
        ``iat``, ``exp`` are the W62 convention).
        """
        ...

    def verify(self, token: str, *, now: int) -> dict[str, Any]:
        """Verify the signature and expiration; return the payload.

        Raises:
            ``InvalidTokenError`` — token is malformed, signature
                does not verify, or base64 decoding fails.
            ``ExpiredTokenError`` — payload ``exp`` is set and ``now``
                is greater than or equal to ``exp``.
        """
        ...


JwtSigner = JwtSignerPort


__all__ = ["JwtSigner", "JwtSignerPort"]
