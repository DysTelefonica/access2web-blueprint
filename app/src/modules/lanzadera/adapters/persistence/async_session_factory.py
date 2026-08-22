# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W01 (#44)
# DA-1, DA-2 — async SQLAlchemy 2.0 Core session factory for adapters (Phase 4).
"""Async SQLAlchemy session factory for the lanzadera Postgres adapters.

W01 (#44) ships the seam between the platform and the Postgres
connection: an ``AsyncEngine`` and a callable ``AsyncSessionFactoryPort``
that the persistence adapters (Phase 4) take as a constructor
dependency. The factory sets the session ``search_path`` to the
canonical schema (``lanzadera``, plus ``public``) so the adapters can
write fully-qualified column queries without per-statement schema
qualifiers.

The factory is intentionally pure — it does not import any domain
Protocol, does not open a connection at construction time (the
``AsyncEngine`` does it lazily on first use), and does not own a global
state. Every consumer acquires a session via ``async_session_factory(db_url)``
and passes the resulting object into the adapters (DA-1: dependency
injection; the adapters do not know about a global engine).
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from typing import Protocol

from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.ext.asyncio import (
    AsyncEngine,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)

# Canonical schema the migration tree uses (`migrations/versions/0001_*`).
# Kept here so the runtime and the migration tree cannot drift silently.
SCHEMA = "lanzadera"


class AsyncSessionFactoryError(RuntimeError):
    """Raised when the engine cannot reach Postgres (refused, bad URL, etc.)."""


class AsyncSessionFactoryPort(Protocol):
    """Wiring seam between the platform and the persistence adapters.

    Adapters take ``AsyncSessionFactoryPort`` instead of an engine so the
    adapter stays decoupled from ``sqlalchemy.engine`` and from a
    concrete URL. The protocol is duck-typed: any callable returning a
    new ``AsyncSession`` per call satisfies it.
    """

    def __call__(self) -> AsyncSession:
        """Return a fresh ``AsyncSession``.

        Callers are responsible for closing the session (the helpers
        below are context managers; the adapters use
        ``try/except/finally`` directly).
        """
        ...


@dataclass(frozen=True)
class AsyncSessionFactory:
    """Default ``AsyncSessionFactoryPort`` implementation.

    Holds an ``async_sessionmaker[AsyncSession]`` bound to the configured
    ``AsyncEngine``. The factory is process-wide; running the platform
    in two parallel processes against the same Postgres is out of
    scope for Phase 4.

    ``session_kwargs`` control the per-session Postgres connection
    state: ``search_path`` is the schema-with-fallback constant; the
    Postgres extension setting ``application_name`` is the platform
    identifier so the operations team can correlate long-running
    queries to this service.
    """

    maker: async_sessionmaker[AsyncSession]
    _engine: AsyncEngine

    def __call__(self) -> AsyncSession:
        """Return a brand-new ``AsyncSession``.

        The session is bound to the engine but has not issued any
        queries yet. Callers must ``close()`` it (use ``async with`` in
        long-running flows).
        """
        return self.maker()

    @asynccontextmanager
    async def transaction(self) -> AsyncIterator[AsyncSession]:
        """Yield an ``AsyncSession`` inside an implicit transaction.

        Convenience wrapper so the adapters do not have to know the
        ``async with`` shape. The yielded ``AsyncSession`` is the same
        one the callable returns.
        """
        session = self.maker()
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()


def async_session_factory(
    url: str,
    *,
    application_name: str = "lanzadera",
) -> tuple[AsyncEngine, AsyncSessionFactoryPort]:
    """Build the ``AsyncEngine`` + ``AsyncSessionFactoryPort`` pair.

    The factory sets ``search_path`` once per session via
    execution-options so every connection sees ``lanzadera`` first and
    falls back to ``public``. Postgres ENUMs in other schemas (e.g.
    the seed) are reachable through ``public``.
    """

    def _connect_args() -> dict[str, dict[str, str]]:
        return {
            "server_settings": {
                "application_name": application_name,
                "search_path": f"{SCHEMA},public",
            },
        }

    try:
        engine = create_async_engine(
            url,
            future=True,
            connect_args=_connect_args(),
        )
    except SQLAlchemyError as exc:
        raise AsyncSessionFactoryError(
            f"could not create AsyncEngine for url={url!r}: {exc}"
        ) from exc

    maker = async_sessionmaker(
        engine,
        expire_on_commit=False,
        class_=AsyncSession,
    )

    factory = AsyncSessionFactory(maker=maker, _engine=engine)
    return engine, factory


__all__ = [
    "AsyncSessionFactory",
    "AsyncSessionFactoryError",
    "AsyncSessionFactoryPort",
    "SCHEMA",
    "async_session_factory",
]
