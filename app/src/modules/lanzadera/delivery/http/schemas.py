# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#543)
"""HTTP-bound Pydantic schemas for the JSON routers (issue #774).

The W62 auth flow and the W57 admin flow own their respective schemas
in the same module so delivery keeps a single import surface and stays
free of domain-layer types in handler signatures. ``CreateUserRequest``
is the body shape that ``POST /admin/users`` advertises in the OpenAPI
spec; the handler binds the request body through it so a JSON
payload with the three required fields is the contract, instead of
the legacy ``*, email: str, name: str, national_id: str`` signature
(which silently coerced every field into a query parameter and
rejected real JSON bodies with ``422 Field required [query, email]``).
"""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field

# Lightweight email regex — pydantic's ``EmailStr`` is not an option
# here because it pulls the ``email-validator`` extra into the
# runtime, which the lockfile does not pin. The validation stays
# strict enough to reject obviously-malformed payloads without
# adding a dependency for what is a single character class.
_EMAIL_PATTERN: str = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"


class CreateUserRequest(BaseModel):
    """Body schema for ``POST /admin/users`` (issue #774).

    Mirrors the column constraints in the W51 migration: email is a
    canonical ``str`` matching an ``@``-shaped pattern (the CITEXT
    column stores it lower-cased by the handler), name is a short
    display string (the migration stores it as ``text NOT NULL``),
    and ``national_id`` is the opaque PGP-encrypted ciphertext owner
    (the route never decrypts it; only the platform user CLI does).

    Strip / lower-case happen in the handler so the validator stays
    pure data shape — ``str.strip`` would also affect whitespace
    inside the string, which may be a feature for some clients and
    a bug for others. Keeping the transform out of the schema pins
    the contract to "what the client sends", not "what the client
    sends *after* the server rewrites it".
    """

    model_config = ConfigDict(extra="forbid")

    email: str = Field(min_length=3, max_length=320, pattern=_EMAIL_PATTERN)
    name: str = Field(min_length=1, max_length=200)
    national_id: str = Field(min_length=1, max_length=32)


__all__ = ["CreateUserRequest"]
