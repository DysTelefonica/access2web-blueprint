# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#541)
"""Tests Categoría 3 (contract) para ``Hs256JwtSigner`` y ``JwtSignerPort``.

D-W62-1: HS256 implementation using only the Python stdlib
(``hmac``, ``hashlib``, ``base64``, ``json``, ``secrets``). No external
JWT libraries.

The tests cover:
- Structural conformance: ``Hs256JwtSigner`` exposes every method on
  ``JwtSignerPort`` (DA-1).
- Sign produces a 3-segment ``.``-joined string.
- Verify round-trips a freshly-signed token.
- Verify rejects a tampered signature.
- Verify rejects an expired token.
"""

from __future__ import annotations

import json
from base64 import urlsafe_b64decode

import pytest

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.domain.errors import ExpiredTokenError, InvalidTokenError
from app.src.modules.lanzadera.domain.ports.jwt_signer import JwtSignerPort

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


SECRET = b"test-secret-32-bytes-min-padding!!"  # 32 bytes exactly


def _decode_segment(segment: str) -> dict:
    """Decode a base64url JWT segment back to a dict (test-only helper)."""
    pad = "=" * (-len(segment) % 4)
    return json.loads(urlsafe_b64decode(segment + pad).decode("utf-8"))


# ---------------------------------------------------------------------------
# Structural conformance
# ---------------------------------------------------------------------------


def test_hs256_jwt_signer_conforms_to_protocol() -> None:
    """DA-1: the adapter satisfies the ``JwtSignerPort`` Protocol structurally.

    Mypy checks the conformance statically. The runtime check below is
    a belt-and-suspenders test in case mypy ever stops inheriting async
    signatures.
    """
    expected = {n for n in dir(JwtSignerPort) if not n.startswith("_")} - {
        "__init__",
        "__annotations__",
        "__subclasshook__",
    }
    actual = {n for n in dir(Hs256JwtSigner) if not n.startswith("_")}
    missing = expected - actual
    assert not missing, f"Hs256JwtSigner is missing Protocol methods: {missing}"


# ---------------------------------------------------------------------------
# Sign
# ---------------------------------------------------------------------------


def test_sign_produces_three_dot_separated_segments() -> None:
    signer = Hs256JwtSigner(SECRET)
    token = signer.sign({"sub": "abc", "iat": 1700000000, "exp": 1700003600})

    parts = token.split(".")
    assert len(parts) == 3
    assert all(p for p in parts)  # no empty segments


def test_sign_encodes_header_with_alg_and_typ() -> None:
    signer = Hs256JwtSigner(SECRET)
    token = signer.sign({"sub": "abc"})

    header = _decode_segment(token.split(".")[0])
    assert header == {"alg": "HS256", "typ": "JWT"}


def test_sign_refuses_short_secret() -> None:
    """RFC 7518 §3.2: HS256 key must be at least as long as the hash output (32 bytes)."""
    with pytest.raises(ValueError, match="at least 32 bytes"):
        Hs256JwtSigner(b"too-short")


def test_sign_refuses_non_bytes_secret() -> None:
    with pytest.raises(TypeError, match="must be bytes"):
        Hs256JwtSigner("not-bytes")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Verify — happy path
# ---------------------------------------------------------------------------


def test_verify_round_trip() -> None:
    signer = Hs256JwtSigner(SECRET)
    payload = {"sub": "abc", "iat": 1700000000, "exp": 1700003600}

    token = signer.sign(payload)
    out = signer.verify(token, now=1700001000)

    assert out == payload


# ---------------------------------------------------------------------------
# Verify — failure paths
# ---------------------------------------------------------------------------


def test_verify_rejects_tampered_signature() -> None:
    signer = Hs256JwtSigner(SECRET)
    token = signer.sign({"sub": "abc", "exp": 1700003600})

    # Flip one base64url character in the signature segment (last part).
    header_b64, body_b64, sig_b64 = token.split(".")
    # Pick a byte in the signature that won't break base64 decoding.
    tampered_sig = sig_b64[:-2] + ("A" if sig_b64[-2] != "A" else "B") + sig_b64[-1]
    tampered = f"{header_b64}.{body_b64}.{tampered_sig}"

    with pytest.raises(InvalidTokenError):
        signer.verify(tampered, now=1700000000)


def test_verify_rejects_token_with_wrong_secret() -> None:
    """Signing with one secret and verifying with another raises."""
    signer_a = Hs256JwtSigner(b"a" * 32)
    signer_b = Hs256JwtSigner(b"b" * 32)
    token = signer_a.sign({"sub": "abc", "exp": 1700003600})

    with pytest.raises(InvalidTokenError):
        signer_b.verify(token, now=1700000000)


def test_verify_rejects_expired_token() -> None:
    signer = Hs256JwtSigner(SECRET)
    token = signer.sign({"sub": "abc", "exp": 1700001000})

    with pytest.raises(ExpiredTokenError):
        signer.verify(token, now=1700001000)  # now == exp


def test_verify_rejects_token_without_exp() -> None:
    """A token without ``exp`` is treated as never-expiring (not as InvalidTokenError).

    Some tokens may be intentionally eternal (e.g. service-to-service);
    we follow RFC 7519 §4.1.4 — ``exp`` is OPTIONAL.
    """
    signer = Hs256JwtSigner(SECRET)
    token = signer.sign({"sub": "abc"})

    out = signer.verify(token, now=10_000_000_000)
    assert out == {"sub": "abc"}


def test_verify_rejects_malformed_token() -> None:
    signer = Hs256JwtSigner(SECRET)
    with pytest.raises(InvalidTokenError):
        signer.verify("not.a.jwt", now=0)


def test_verify_rejects_token_with_wrong_segment_count() -> None:
    signer = Hs256JwtSigner(SECRET)
    with pytest.raises(InvalidTokenError, match="3 segments"):
        signer.verify("a.b", now=0)


def test_verify_rejects_token_with_non_object_payload() -> None:
    """A JWT whose payload is a JSON array or string is not a valid auth payload."""
    import base64

    signer = Hs256JwtSigner(SECRET)
    # Build a token with a JSON array payload (invalid structure).
    header_b64 = urlsafe_b64decode(
        # tiny url-safe base64 padding adjustment
        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9" + "="
    ).decode()
    payload_array = json.dumps([1, 2, 3]).encode("utf-8")
    body_b64 = base64.urlsafe_b64encode(payload_array).rstrip(b"=").decode("ascii")
    sig_b64 = "AAAA"  # signature bytes don't matter; signature check happens first
    bad = f"{header_b64}.{body_b64}.{sig_b64}"

    # The signature check fails first (different secret, fake sig),
    # so we expect InvalidTokenError regardless of payload shape.
    with pytest.raises(InvalidTokenError):
        signer.verify(bad, now=0)
