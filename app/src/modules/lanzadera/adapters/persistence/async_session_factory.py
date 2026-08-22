"""Postgres session factory (lanzadera-mvp, AD1, issue #44).

Wraps ``sqlalchemy.ext.asyncio.async_sessionmaker`` with the schema-aware
reflection so every ``AsyncSession`` is bound to ``lanzadera.*`` by
default. The factory is the only seam between the rest of the platform
and the database driver: the composition root (issue #49) injects
one factory per process and the use-case layers consume the resulting
sessions.

The factory accepts either a ``DatabaseURL`` (``postgresql+asyncpg://...``)
or a pre-built ``AsyncEngine``. The latter path is what tests use (the
test fixtures build a real engine against a freshly-provisioned
``CREATE DATABASE`` instance and pass it through), the former is what
production uses (the env var ``DATABASE_URL`` is parsed by the
composition root and passed as-is).
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from typing import Protocol

from sqlalchemy.engine.url import URL, make_url
from sqlalchemy.ext.asyncio import (
    AsyncEngine,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)

# Every session this factory opens sees the ``lanzadera`` schema by
# default. This default schema is enforced by the use-case SQL (``SET
# search_path TO lanzadera, public``) so the production migrations in
# 0001+ can target the named schema without per-table qualifiers.
DEFAULT_SCHEMA = "lanzadera"


class AsyncSessionFactoryPort(Protocol):
    """Structural Protocol — the use-case layer's view of a session."""

    def __call__(self) -> AsyncSession: ...


def async_session_factory(
    url: str | URL,
    *,
    schema_search_path: tuple[str, ...] = (DEFAULT_SCHEMA, "public"),
    echo: bool = False,
    pool_size: int = 5,
    pool_max_overflow: int = 10,
) -> tuple[AsyncEngine, AsyncSessionFactoryPort]:
    """Build an async engine + factory whose sessions target ``lanzadera``.

    The factory is a closure over the engine. Callers (composition root)
    get back a tuple of (engine, factory); the factory is what the
    use-case layer depends on via ``AsyncSessionFactoryPort`` (a Protocol
    that matches ``AsyncSessionmaker.__call__`` structurally).

    The ``schema_search_path`` defaults to ``("lanzadera", "public")``
    which matches what Alembic emits in the migration 0001 (the
    ``lanzadera`` schema for application tables, ``public`` for the
    ``gen_random_uuid()`` extension function). Override only when
    running the integration tests against a temporary schema.
    """
    parsed_url: URL = url if isinstance(url, URL) else make_url(url)
    engine = create_async_engine(
        parsed_url,
        echo=echo,
        pool_size=pool_size,
        max_overflow=pool_max_overflow,
        connect_args={"server_settings": {"search_path": ", ".join(schema_search_path)}},
    )
    session_factory: AsyncSessionFactoryPort = async_sessionmaker(
        engine, expire_on_commit=False, class_=AsyncSession
    )
    return engine, session_factory


async def with_session(
    factory: AsyncSessionFactoryPort,
) -> AsyncIterator[AsyncSession]:
    """Yield a session, commit on clean exit, rollback on exception.

    Convenience wrapper for the application-layer use-cases. Callers
    ``async with with_session(factory) as session:`` then read or write
    through the session, returning the value commits the transaction.
    """
    session = factory()
    try:
        yield session
    except Exception:
        await session.rollback()
        raise
    else:
        await session.commit()
    finally:
        await session.close()
