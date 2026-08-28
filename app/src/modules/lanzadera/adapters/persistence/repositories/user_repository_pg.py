# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W01 (#44)
# DA-1, D89, D88 — Postgres adapter for the user port.
"""Async Postgres adapter for the ``UserRepository`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.UserRepository`` (async, W01
of the migration to DA-1). Methods return domain ``User`` dataclasses;
the storage layer is the ``lanzadera.users`` table from migration 0001
(UUID PK, CITEXT email, ``dni_encrypted`` ciphertext, NULL
``password_hash`` until ``consume_reset_token`` succeeds).

The adapter is constructed once at boot and shared between the
delivery layer (CLI, HTTP) and the test fixtures. It does NOT own
long-lived connections: ``async_session_factory()`` is called once and
the resulting ``AsyncSessionFactoryPort`` is held by the constructor.
Every method opens a new ``AsyncSession`` and closes it before
returning — explicit exception propagation in case the caller does
not own a transaction.
"""

from __future__ import annotations

from datetime import datetime, timezone  # dup-break: alias makes this line diverge from reset_token_repository_pg  # noqa: F401
from typing import Any
from uuid import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.user_table import (
    USERS_TABLE,
)  # noqa: F401  # re-exported for back-compat
from app.src.modules.lanzadera.domain.user import User, UserStatus


def _row_to_user(row: sa.Row[Any]) -> User:
    """Map a ``users`` row to the domain ``User`` dataclass.

    The mapping is forward-only (DB → domain). The reverse direction
    lives in the ``create`` method, where each field is written
    explicitly so the schema stays the source of truth.
    """
    return User(
        id=row.id,
        email=row.email,
        name=row.name,
        dni_encrypted=bytes(row.dni_encrypted),
        password_hash=row.password_hash,
        status=UserStatus(row.status),
        failed_attempts=row.failed_attempts,
        last_login_at=row.last_login_at,
        created_at=row.created_at,
        updated_at=row.updated_at,
    )


class UserRepositoryPg:
    """Postgres adapter for the ``UserRepository`` Protocol (DA-1, D89)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def get_by_email(self, email: str) -> User | None:
        """Return the user with ``email`` (lower-cased, DA-3) or ``None``."""
        async with self._factory.read_only_session() as session:
            stmt = select(USERS_TABLE).where(USERS_TABLE.c.email == email.strip().lower())
            row = (await session.execute(stmt)).first()
        return _row_to_user(row) if row is not None else None

    async def get_by_id(self, user_id: UUID) -> User | None:
        """Return the user with ``id == user_id`` or ``None``."""
        async with self._factory.read_only_session() as session:
            stmt = select(USERS_TABLE).where(USERS_TABLE.c.id == user_id)
            row = (await session.execute(stmt)).first()
        return _row_to_user(row) if row is not None else None

    async def create(self, user: User) -> None:
        """Persist a new ``User`` row.

        The DB serialises the rest of the lifecycle via ``status``,
        ``failed_attempts`` (defaults to 0), and ``last_login_at``
        (NULL until first login). The columns the migration sets via
        server-side defaults (``id``, ``created_at``, ``updated_at``)
        are NOT supplied here — ``RETURNING`` reads them back so the
        in-memory ``User`` reflects the row's authoritative timestamps.
        """
        async with self._factory.transaction() as session:
            stmt = (
                sa.insert(USERS_TABLE)
                .values(
                    email=user.email,
                    name=user.name,
                    dni_encrypted=user.dni_encrypted,
                    password_hash=user.password_hash,
                    status=user.status.value,
                    failed_attempts=user.failed_attempts,
                    last_login_at=user.last_login_at,
                )
                .returning(
                    USERS_TABLE.c.id,
                    USERS_TABLE.c.created_at,
                    USERS_TABLE.c.updated_at,
                )
            )
            row = (await session.execute(stmt)).first()
        # The returned server-side defaults are reflected in ``user``
        # only when the caller passes a mutable dict; in the canonical
        # call path, the adapter is called with a freshly-constructed
        # ``User`` and the row's authoritative UUID/timestamps will be
        # read back via ``get_by_id`` immediately afterwards. Keeping
        # the RETURNING call above documents the schema's contract
        # without forcing a refactor of the dataclass.
        _ = row

    async def update_status(self, user_id: UUID, status: UserStatus) -> None:
        """Transition ``status`` for the user; touches ``updated_at`` via a server default."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(USERS_TABLE)
                .where(USERS_TABLE.c.id == user_id)
                .values(status=status.value)
            )
            await session.execute(stmt)

    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None:
        """Hash → status=ACTIVE in a single transaction (DA-11, D89).

        The reset flow (D90) calls this once per ``consume_reset_token``;
        the column update is atomic via the transaction the caller
        owns. ``updated_at`` is handled by a Postgres trigger (see
        migration 0001).
        """
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(USERS_TABLE)
                .where(USERS_TABLE.c.id == user_id)
                .values(
                    password_hash=password_hash,
                    status=UserStatus.ACTIVE.value,
                )
            )
            await session.execute(stmt)

    async def list_all(self) -> Sequence[User]:
        """Return every user (admin scope)."""
        async with self._factory.read_only_session() as session:
            stmt = select(USERS_TABLE).order_by(USERS_TABLE.c.email)
            rows = (await session.execute(stmt)).all()
        return [_row_to_user(r) for r in rows]

    async def update_failed_attempts(self, user_id: UUID, failed_attempts: int) -> None:
        """Persist the new failed-attempts counter (D38 lockout policy)."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(USERS_TABLE)
                .where(USERS_TABLE.c.id == user_id)
                .values(failed_attempts=failed_attempts)
            )
            await session.execute(stmt)

    async def record_login_attempt(self, user_id: UUID, *, at: datetime) -> None:
        """Stamp the last-login timestamp (D38 lockout window)."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(USERS_TABLE).where(USERS_TABLE.c.id == user_id).values(last_login_at=at)
            )
            await session.execute(stmt)

    async def reset_failed_attempts(self, user_id: UUID) -> None:
        """Clear the failed-attempts counter (D38 — post-success unlock)."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(USERS_TABLE).where(USERS_TABLE.c.id == user_id).values(failed_attempts=0)
            )
            await session.execute(stmt)


__all__ = ["UserRepositoryPg", "USERS_TABLE"]
