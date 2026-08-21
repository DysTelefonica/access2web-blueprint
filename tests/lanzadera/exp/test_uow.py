"""Unit tests for the Expedientes F04 foundations.

This file covers the two foundations F04 ships in ``app/src/modules/expedientes/``:

1. :class:`UnitOfWork` — the transactional context manager (DA-11
   contract: commit on success, rollback on exception, close on either).
2. :class:`SessionFactory` Protocol — the structural contract the
   composition root injects.

The migration itself (``app/migrations/versions/0002_expedientes_schema.py``)
is covered by ``tests/lanzadera/migrations/test_migration_0001_downgrade.py``
in a follow-up WU: writing a migration integration test for
``0002`` requires a live PostgreSQL in the CI env (the ``0001`` pattern
already does that — see ``tests/lanzadera/migrations/test_migration_0001.py``,
gated on ``APAP_INTEGRATION_ENABLED=1``). For F04 we ship the unit-level
UoW tests so the verticals (C01..R07, M01..M04) can build against a
verified ``UnitOfWork`` contract.

The UoW is the foundation that DA-11 (``AuditLogPort.append MUST run in
the same SQLAlchemy session as the mutating case-of-use so a failed
insert causes ROLLBACK of the auth mutation``) relies on. The verticals
that follow will not introduce a new transactional helper; they will
import this one and call ``with UnitOfWork(factory) as uow:``.
"""

from __future__ import annotations

from typing import Self
from uuid import uuid4

import pytest

from app.src.modules.expedientes.shared.uow import SessionFactory, UnitOfWork


class _FakeSession:
    """In-memory SQLAlchemy-shaped session for the UoW tests.

    The real :class:`UnitOfWork` only calls ``commit`` / ``rollback`` /
    ``close`` on the session, plus exposes the session as ``uow.session``
    for the application code to use. The application code (verticals)
    uses ``add`` / ``delete`` / ``query`` / ``get``; this fake tracks
    those operations so tests can assert the application layer wired
    the right objects into the session.
    """

    def __init__(self) -> None:
        self.committed: bool = False
        self.rolled_back: bool = False
        self.closed: bool = False
        self.added: list[object] = []
        self.deleted: list[object] = []
        self.queried: list[object] = []

    def commit(self) -> None:
        self.committed = True

    def rollback(self) -> None:
        self.rolled_back = True

    def close(self) -> None:
        self.closed = True

    def execute(self, statement: object) -> object:  # noqa: ARG002
        return None

    def merge(self, instance: object) -> object:
        return instance

    def add(self, instance: object) -> None:
        self.added.append(instance)

    def delete(self, instance: object) -> None:
        self.deleted.append(instance)

    def query(self, *entities: object) -> Self:  # returns self for chaining
        self.queried.extend(entities)
        return self

    def get(self, entity: object, ident: object) -> object | None:
        return None


class _FakeFactory:
    """Session factory that hands out a fresh ``_FakeSession`` per call."""

    def __init__(self) -> None:
        self.sessions: list[_FakeSession] = []

    def __call__(self) -> _FakeSession:
        session = _FakeSession()
        self.sessions.append(session)
        return session


def test_uow_commits_on_clean_exit() -> None:
    factory = _FakeFactory()
    with UnitOfWork(factory) as session:
        session.add("row1")
    assert len(factory.sessions) == 1
    assert factory.sessions[0].committed is True
    assert factory.sessions[0].rolled_back is False
    assert factory.sessions[0].closed is True


def test_uow_rolls_back_on_exception() -> None:
    factory = _FakeFactory()
    with pytest.raises(RuntimeError, match="boom"):
        with UnitOfWork(factory) as session:
            session.add("row1")
            raise RuntimeError("boom")
    assert factory.sessions[0].committed is False
    assert factory.sessions[0].rolled_back is True
    assert factory.sessions[0].closed is True
    # And the application-layer add() is still in the session — the
    # ROLLBACK is the database's job, not the Python in-memory store's.
    assert factory.sessions[0].added == ["row1"]


def test_uow_session_access_outside_block_raises() -> None:
    """DA-11: a session that lives outside its transaction is unsafe; the
    UoW MUST refuse to expose it after ``__exit__``."""
    factory = _FakeFactory()
    uow = UnitOfWork(factory)
    # Not yet entered — the ``with`` block has not bound the session.
    with pytest.raises(RuntimeError, match="outside the 'with' block"):
        _ = uow.session  # noqa: B018 — accessing the property raises

    # And the same guard fires AFTER the block: the session reference
    # is invalidated by ``__exit__``.
    with uow:
        pass
    with pytest.raises(RuntimeError, match="outside the 'with' block"):
        _ = uow.session  # noqa: B018


def test_uow_factory_protocol_is_structural() -> None:
    """The :class:`SessionFactory` Protocol is structural — any callable that
    returns an object with ``commit`` / ``rollback`` / ``close`` satisfies
    it. The composition root injects a SQLAlchemy ``sessionmaker``; the
    tests inject ``_FakeFactory``. The ``isinstance`` check is irrelevant
    here (we exercise the runtime contract, not the type check).
    """
    factory: SessionFactory = _FakeFactory()  # type: ignore[assignment]
    with UnitOfWork(factory) as session:
        assert session.add("row1") is None
    assert factory.sessions[0].committed is True


def test_uow_each_call_gets_fresh_session() -> None:
    """The factory is called exactly once per ``with`` block, and the
    session is unique to the block. Two consecutive ``UnitOfWork``s
    MUST NOT share state (DA-11 atomicity)."""
    factory = _FakeFactory()
    with UnitOfWork(factory):
        pass
    with UnitOfWork(factory):
        pass
    assert len(factory.sessions) == 2
    assert factory.sessions[0] is not factory.sessions[1]


def test_uow_exception_propagates_after_close() -> None:
    """The UoW MUST NOT swallow exceptions. DA-11 requires the application
    layer to observe the failure so it can map to a 4xx/5xx response."""
    factory = _FakeFactory()
    sentinel = ValueError("sentinel")
    with pytest.raises(ValueError) as exc_info:
        with UnitOfWork(factory):
            raise sentinel
    assert exc_info.value is sentinel
    assert factory.sessions[0].closed is True


def test_uow_supports_multiple_mutations_in_one_block() -> None:
    """A single UoW MUST support many ``add`` / ``delete`` calls and commit
    them atomically (single transaction). The verticals' case-of-use
    code typically adds the aggregate header, then the child rows,
    then the read-model, in one block; this test pins the
    "one block, one transaction" contract."""
    factory = _FakeFactory()
    record = type("Aggregate", (), {"id": uuid4()})()
    children = [type("Child", (), {"id": uuid4()})() for _ in range(3)]
    with UnitOfWork(factory) as session:
        session.add(record)
        for child in children:
            session.add(child)
    assert factory.sessions[0].added == [record, *children]
    assert factory.sessions[0].committed is True


def test_uow_session_attribute_is_bound_to_entered_session() -> None:
    """The session reference held during one block is unique to that block.

    ``__enter__`` returns the session, so the application's typical use
    is ``with UnitOfWork(factory) as session: ...``. The test here
    captures the session object from each block and asserts the two
    are distinct objects.
    """
    factory = _FakeFactory()
    with UnitOfWork(factory) as first:
        first_ref = first
    with UnitOfWork(factory) as second:
        second_ref = second
    assert first_ref is not second_ref
    assert first_ref is factory.sessions[0]
    assert second_ref is factory.sessions[1]


def test_uow_session_close_even_on_generator_error() -> None:
    """``__exit__`` MUST close the session regardless of whether commit or
    rollback raised. The factory wrapper might raise (a network blip on
    commit, say) and a leaked session is a connection-pool leak."""
    factory = _FakeFactory()

    class _RaisingSession(_FakeSession):
        def commit(self) -> None:
            raise OSError("commit failed")

    def raising_factory() -> _RaisingSession:
        factory.sessions.append(_RaisingSession())  # type: ignore[arg-type]
        return factory.sessions[-1]  # type: ignore[return-value]

    with pytest.raises(IOError, match="commit failed"):
        with UnitOfWork(raising_factory) as session:  # type: ignore[arg-type]
            session.add("row1")
    # The session was created but the commit failed; the UoW still
    # closes it in the ``finally`` of ``__exit__``.
    assert factory.sessions[0].closed is True
