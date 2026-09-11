"""Strict TDD — Hito collection in Expediente aggregate (F03, issue #223).

RED: tests first, then add_hito/remove_hito to Expediente.
"""

from __future__ import annotations

from datetime import UTC, date, datetime
from uuid import uuid4

import pytest

from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.domain.hito import Hito


def _d(year: int, month: int, day: int) -> date:
    return date(year, month, day)


def _dt(year: int, month: int, day: int) -> datetime:
    return datetime(year, month, day, 12, 0, 0, tzinfo=UTC)


def _exp_am(
    *,
    id: uuid4 | None = None,
    estado: ExpedienteEstado = ExpedienteEstado.BORRADOR,
) -> Expediente:
    return Expediente(
        id=id or uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=estado,
        version=1,
        created_at=_dt(2026, 9, 1),
        updated_at=_dt(2026, 9, 1),
    )


def _hito(
    *,
    id: uuid4 | None = None,
    id_expediente: uuid4,
    fecha_hito: date = _d(2026, 10, 1),
    garantia_fecha_fin: date | None = _d(2027, 10, 2),
) -> Hito:
    return Hito(
        id=id or uuid4(),
        id_expediente=id_expediente,
        fecha_hito=fecha_hito,
        garantia_fecha_fin=garantia_fecha_fin,
        estado=ExpedienteEstado.BORRADOR,
        created_at=_dt(2026, 9, 1),
        updated_at=_dt(2026, 9, 1),
    )


# ---------------------------------------------------------------------------
# Hito collection — initial state
# ---------------------------------------------------------------------------

class TestExpedienteHitosInitialState:
    """Expediente starts with an empty hitos list."""

    def test_hitos_is_empty_on_construction(self) -> None:
        exp = _exp_am()
        assert list(exp.hitos) == []

    def test_hitos_is_readonly_view(self) -> None:
        """The hitos property returns an immutable view."""
        exp = _exp_am()
        with pytest.raises(AttributeError):
            exp.hitos.append("not allowed")  # type: ignore


# ---------------------------------------------------------------------------
# add_hito — happy path
# ---------------------------------------------------------------------------

class TestAddHitoHappyPath:
    """Hito added to the correct Expediente is stored in the collection."""

    def test_add_hito_stores_it(self) -> None:
        exp = _exp_am()
        hito = _hito(id_expediente=exp.id)
        exp.add_hito(hito)
        assert list(exp.hitos) == [hito]

    def test_add_multiple_hitos(self) -> None:
        exp = _exp_am()
        h1 = _hito(id_expediente=exp.id, fecha_hito=_d(2026, 10, 1))
        h2 = _hito(id_expediente=exp.id, fecha_hito=_d(2026, 12, 1))
        exp.add_hito(h1)
        exp.add_hito(h2)
        assert list(exp.hitos) == [h1, h2]

    def test_added_hito_references_correct_expediente(self) -> None:
        exp = _exp_am()
        hito = _hito(id_expediente=exp.id)
        exp.add_hito(hito)
        assert hito.id_expediente == exp.id


# ---------------------------------------------------------------------------
# add_hito — invariants
# ---------------------------------------------------------------------------

class TestAddHitoInvariants:
    """add_hito enforces aggregate invariants."""

    def test_rejects_hito_with_wrong_expediente_id(self) -> None:
        """CAP-008: a Hito belongs to exactly one Expediente."""
        exp = _exp_am()
        wrong_exp = _exp_am()
        hito = _hito(id_expediente=wrong_exp.id)
        with pytest.raises(ValueError, match="wrong Expediente"):
            exp.add_hito(hito)

    def test_rejects_duplicate_hito_id(self) -> None:
        """A Hito with the same id cannot be added twice."""
        exp = _exp_am()
        hito_id = uuid4()
        h1 = _hito(id=hito_id, id_expediente=exp.id)
        h2 = _hito(id=hito_id, id_expediente=exp.id)
        exp.add_hito(h1)
        with pytest.raises(ValueError, match="already exists"):
            exp.add_hito(h2)


# ---------------------------------------------------------------------------
# remove_hito — happy path
# ---------------------------------------------------------------------------

class TestRemoveHitoHappyPath:
    """Hito removed from the collection is no longer present."""

    def test_remove_existing_hito(self) -> None:
        exp = _exp_am()
        hito = _hito(id_expediente=exp.id)
        exp.add_hito(hito)
        exp.remove_hito(hito.id)
        assert list(exp.hitos) == []

    def test_remove_one_of_multiple(self) -> None:
        exp = _exp_am()
        h1 = _hito(id_expediente=exp.id, fecha_hito=_d(2026, 10, 1))
        h2 = _hito(id_expediente=exp.id, fecha_hito=_d(2026, 12, 1))
        exp.add_hito(h1)
        exp.add_hito(h2)
        exp.remove_hito(h1.id)
        assert list(exp.hitos) == [h2]


# ---------------------------------------------------------------------------
# remove_hito — invariants
# ---------------------------------------------------------------------------

class TestRemoveHitoInvariants:
    """remove_hito enforces aggregate invariants."""

    def test_remove_nonexistent_hito_raises(self) -> None:
        exp = _exp_am()
        with pytest.raises(ValueError, match="not found"):
            exp.remove_hito(uuid4())
