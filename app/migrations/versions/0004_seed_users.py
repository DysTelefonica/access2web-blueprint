# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #620
"""seed users: insert 156 rows from data/fixtures/lanzadera/users.json.

Revision ID: 0004
Revises: 0003
Create Date: 2026-09-16

DA-3, D89: every seeded user gets ``password_hash = NULL`` and
``status = 'password_reset_required'``. The ``legacy_hash`` column
does not exist in this schema — the migration does not write it.

The ``dni`` field is encrypted to ``dni_encrypted`` using the same
AES-256-GCM + Fernet-key layout as ``NationalIdCipher`` in the app
(so downstream code can decrypt with the same ``PLATFORM_SECRET_KEY``).
The encryption is done inline here because ``migrations/`` is driven and
may not import from ``adapters/`` (DA-1)."""

from __future__ import annotations

import base64
import json
import os
from collections.abc import Sequence
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0004"
down_revision: str | None = "0003"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def _encrypt_national_id(dni_plain: str, raw_key: bytes) -> bytes:
    """AES-256-GCM encrypt ``dni_plain`` → nonce(12) || ciphertext || tag(16)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    nonce = os.urandom(12)
    return nonce + AESGCM(raw_key).encrypt(nonce, dni_plain.encode("utf-8"), None)


def _load_key() -> bytes:
    raw = os.environ["PLATFORM_SECRET_KEY"]
    return base64.urlsafe_b64decode(raw)


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data" / "fixtures" / "lanzadera" / "users.json"
    )
    users = json.loads(fixture.read_text())

    bind = op.get_bind()
    count = (
        bind.execute(sa.text(f"SELECT COUNT(*) FROM {SCHEMA}.users"))
        .scalar()
        or 0
    )
    if count >= len(users):
        return

    key = _load_key()
    for user in users:
        dni_encrypted = _encrypt_national_id(user["dni"], key)
        bind.execute(
            sa.text(
                f"""INSERT INTO {SCHEMA}.users
                    (id, email, name, dni_encrypted, password_hash, status,
                     failed_attempts, created_at, updated_at)
                    VALUES
                    (:id, :email, :name, :dni_encrypted, NULL,
                     'password_reset_required'::{SCHEMA}.user_status,
                     0, now(), now())"""
            ),
            {
                "id": user["id"],
                "email": user["email"],
                "name": user["name"],
                "dni_encrypted": dni_encrypted,
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    bind.execute(sa.text(f"DELETE FROM {SCHEMA}.users WHERE email LIKE '%@synthetic.lanzadera'"))
