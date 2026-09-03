# Spec: JWT utility (PR-4)

> Capability del change `w62-auth-flow`. Firma y verificación de tokens JWT con HS256.

## Purpose

El middleware de auth (PR-5) necesita convertir un `session_id` (UUID) en un token opaco firmado que el cliente envía en `Authorization: Bearer <token>`. La forma estándar es JWT con HS256.

Decisión D-W62-1: implementación stdlib pura (`hmac`, `hashlib`, `base64`, `secrets`). Sin `pyjwt` ni `authlib` para minimizar superficie de seguridad.

## Requirements

### Requirement: HS256 sign

The system SHALL sign a payload dict with HS256 via
``hmac.new(secret, msg, hashlib.sha256).digest()`` and produce a JWT
string of the form ``base64url(header).base64url(payload).base64url(sig)``.

#### Scenario: Sign a valid payload

- GIVEN a secret ``b"test-secret-32-bytes-min-padding!!"`` and a payload ``{"sub": "abc", "iat": 1700000000, "exp": 1700003600}``
- WHEN ``sign(payload)`` is called
- THEN a 3-segment ``.``-joined string is returned
- AND the first segment decodes to ``{"alg":"HS256","typ":"JWT"}``
- AND the second segment decodes to the input payload
- AND the third segment is a 32-byte HMAC-SHA256 digest b64url-encoded

### Requirement: HS256 verify

The system SHALL verify a token's signature by recomputing
``HMAC-SHA256(secret, header_b64.payload_b64)`` and constant-time
comparing against the provided signature.

#### Scenario: Verify a freshly-signed token

- GIVEN a token signed by the same signer
- WHEN ``verify(token)`` is called
- THEN the payload dict is returned

#### Scenario: Verify rejects a tampered signature

- GIVEN a token whose signature byte was flipped
- WHEN ``verify(token)`` is called
- THEN ``InvalidTokenError`` is raised

#### Scenario: Verify rejects an expired token

- GIVEN a token whose payload ``exp`` is in the past
- WHEN ``verify(token, now=<exp+1>)`` is called
- THEN ``ExpiredTokenError`` is raised

#### Scenario: Verify rejects a malformed token

- GIVEN ``"not.a.jwt"`` (only 3 segments, but signature tampered)
- WHEN ``verify("not.a.jwt")`` is called
- THEN ``InvalidTokenError`` is raised

### Requirement: Port surface

The system SHALL expose the contract via ``JwtSignerPort`` so the
production code path uses the ``hmac`` stdlib adapter and the test path
uses ``FakeJwtSigner``.

```python
class JwtSignerPort(Protocol):
    def sign(self, payload: dict[str, Any]) -> str: ...
    def verify(self, token: str, *, now: int) -> dict[str, Any]: ...
```

## File-surface contract

| File | Action | Notes |
|---|---|---|
| `app/src/modules/lanzadera/domain/ports/jwt_signer.py` | NEW | `JwtSignerPort` Protocol + `JwtSigner = JwtSignerPort` alias. |
| `app/src/modules/lanzadera/domain/ports/_imports.py` | MODIFY | add `Any` to the prelude if not present (Protocol + Sequence + UUID + Any). |
| `app/src/modules/lanzadera/adapters/crypto/jwt.py` | NEW | `Hs256JwtSigner` class implementing `JwtSignerPort` with `__init__(secret: bytes)`. |
| `tests/lanzadera/_fakes.py` | MODIFY | add `FakeJwtSigner` with a configurable secret + `sign_calls`/`verify_calls` for test introspection. |
| `tests/lanzadera/adapters/crypto/test_jwt_signer.py` | NEW | Categoría 3 — 6 test cases. |

## Tests Cat 3 — `tests/lanzadera/adapters/crypto/test_jwt_signer.py`

1. `test_hs256_jwt_signer_conforms_to_protocol` — `dir(JwtSignerPort) - dir(Hs256JwtSigner)` empty.
2. `test_sign_produces_three_dot_separated_segments` — split on `.`, assert len == 3.
3. `test_sign_encodes_header_with_alg_and_typ` — b64url-decode first segment, assert `{"alg":"HS256","typ":"JWT"}`.
4. `test_verify_round_trip` — sign a payload, verify it, assert equal payload.
5. `test_verify_rejects_tampered_signature` — flip one byte of the signature, assert `InvalidTokenError`.
6. `test_verify_rejects_expired_token` — sign with `exp = now - 1`, verify with `now`, assert `ExpiredTokenError`.

## Stdlib implementation reference

```python
import base64
import hashlib
import hmac
import json
import time

def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")

def _b64url_decode(s: str) -> bytes:
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)

class Hs256JwtSigner:
    def __init__(self, secret: bytes) -> None:
        if len(secret) < 32:
            raise ValueError("HS256 secret must be at least 32 bytes")
        self._secret = secret

    def sign(self, payload: dict[str, Any]) -> str:
        header = _b64url_encode(json.dumps({"alg": "HS256", "typ": "JWT"}, separators=(",", ":")).encode())
        body = _b64url_encode(json.dumps(payload, separators=(",", ":")).encode())
        msg = f"{header}.{body}".encode()
        sig = hmac.new(self._secret, msg, hashlib.sha256).digest()
        return f"{header}.{body}.{_b64url_encode(sig)}"

    def verify(self, token: str, *, now: int) -> dict[str, Any]:
        parts = token.split(".")
        if len(parts) != 3:
            raise InvalidTokenError("token must have 3 segments")
        header_b64, body_b64, sig_b64 = parts
        msg = f"{header_b64}.{body_b64}".encode()
        expected_sig = hmac.new(self._secret, msg, hashlib.sha256).digest()
        actual_sig = _b64url_decode(sig_b64)
        if not hmac.compare_digest(expected_sig, actual_sig):
            raise InvalidTokenError("signature mismatch")
        payload = json.loads(_b64url_decode(body_b64))
        if "exp" in payload and now >= payload["exp"]:
            raise ExpiredTokenError(f"token expired at {payload['exp']}")
        return payload
```

## Decisiones cerradas

- **`compare_digest`** para comparación constant-time (no `==`).
- **Padding** con `=` hasta múltiplo de 4 antes de `b64decode` (RFC 7515).
- **`exp` está en payload**, no en header — el middleware (PR-5) lo verifica.
- **Sin `iat`/`nbf` checks** — sólo `exp`. Si en el futuro hace falta, se añaden sin breaking change.

## Container integration (NO)

Este PR NO toca `LanzaderaContainer`. La integración del signer viene en PR-6 (`auth_routes`) cuando las rutas necesitan firmarlo. Aquí sólo se construye el adapter + port + fake + tests Cat 3.

## BASELINE updates

`Hs256JwtSigner` debería quedar bajo 100 mutation sites (es código pequeño). Si excede, agregar al BASELINE con `target_date="2027-02-13"`.

`FakeJwtSigner` también es código de tests — bajo mutation sites ceiling normalmente.

## Verificación

```bash
# Sólo los nuevos tests
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/adapters/crypto/test_jwt_signer.py -v
# → 6 passed

# Suite completa
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app --cov=app --cov-report=term
# → 713 + 6 = 719 passed, coverage ≥ 82%

# Gates
ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
/usr/bin/python3.12 -m mypy --explicit-package-bases app/
python3 scripts/check_test_classification.py --root tests/lanzadera
```
