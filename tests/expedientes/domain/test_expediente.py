"""Strict TDD — `Expediente` aggregate root (F02, issue #223).

RED: tests first, then domain entities.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

import pytest

from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=UTC)


# ---------------------------------------------------------------------------
# ExpedienteEstado StrEnum
# ---------------------------------------------------------------------------

class TestExpedienteEstadoStrEnum:
    """Must be a StrEnum so JSON serialisation round-trips."""

    def test_members_are_str(self) -> None:
        for member in ExpedienteEstado:
            assert isinstance(member, str)

    def test_borrador_value(self) -> None:
        assert ExpedienteEstado.BORRADOR.value == "BORRADOR"

    def test_adjudicado_value(self) -> None:
        assert ExpedienteEstado.ADJUDICADO.value == "ADJUDICADO"

    def test_formalizado_value(self) -> None:
        assert ExpedienteEstado.FORMALIZADO.value == "FORMALIZADO"

    def test_archivado_value(self) -> None:
        assert ExpedienteEstado.ARCHIVADO.value == "ARCHIVADO"

    def test_round_trip_from_string(self) -> None:
        assert ExpedienteEstado("BORRADOR") is ExpedienteEstado.BORRADOR
        assert ExpedienteEstado("ADJUDICADO") is ExpedienteEstado.ADJUDICADO
        assert ExpedienteEstado("FORMALIZADO") is ExpedienteEstado.FORMALIZADO
        assert ExpedienteEstado("ARCHIVADO") is ExpedienteEstado.ARCHIVADO


# ---------------------------------------------------------------------------
# ExpedienteTipo StrEnum
# ---------------------------------------------------------------------------

class TestExpedienteTipoStrEnum:
    """Must be a StrEnum matching the Postgres enum literal values."""

    def test_members_are_str(self) -> None:
        for member in ExpedienteTipo:
            assert isinstance(member, str)

    def test_am_value(self) -> None:
        assert ExpedienteTipo.AM.value == "AM"

    def test_lote_value(self) -> None:
        assert ExpedienteTipo.LOTE.value == "LOTE"

    def test_based_value(self) -> None:
        assert ExpedienteTipo.BASED.value == "BASED"

    def test_round_trip_from_string(self) -> None:
        assert ExpedienteTipo("AM") is ExpedienteTipo.AM
        assert ExpedienteTipo("LOTE") is ExpedienteTipo.LOTE
        assert ExpedienteTipo("BASED") is ExpedienteTipo.BASED


# ---------------------------------------------------------------------------
# Expediente construction — happy path
# ---------------------------------------------------------------------------

class TestExpedienteConstruction:
    """Minimum viable `Expediente` carries identity, tipo, estado and version."""

    def test_carries_identity_attributes(self) -> None:
        exp_id = uuid4()
        exp = Expediente(
            id=exp_id,
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert exp.id == exp_id
        assert exp.tipo is ExpedienteTipo.AM
        assert exp.estado is ExpedienteEstado.BORRADOR
        assert exp.version == 1

    def test_parent_is_none_by_default(self) -> None:
        exp = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert exp.id_expediente_padre is None

    def test_is_mutable_for_state_transitions(self) -> None:
        exp = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        exp.estado = ExpedienteEstado.ADJUDICADO
        exp.version = 2
        exp.updated_at = _now()
        assert exp.estado is ExpedienteEstado.ADJUDICADO
        assert exp.version == 2


# ---------------------------------------------------------------------------
# Expediente invariants
# ---------------------------------------------------------------------------

class TestExpedienteInvariants:
    """Domain invariants enforced at construction time."""

    def test_version_must_be_positive(self) -> None:
        with pytest.raises(ValueError, match="version"):
            Expediente(
                id=uuid4(),
                tipo=ExpedienteTipo.AM,
                estado=ExpedienteEstado.BORRADOR,
                version=0,
                created_at=_now(),
                updated_at=_now(),
            )

    def test_version_negative_rejected(self) -> None:
        with pytest.raises(ValueError, match="version"):
            Expediente(
                id=uuid4(),
                tipo=ExpedienteTipo.AM,
                estado=ExpedienteEstado.BORRADOR,
                version=-1,
                created_at=_now(),
                updated_at=_now(),
            )

    def test_lote_requires_parent_id(self) -> None:
        """CAP-006: Lote MUST have a parent reference."""
        with pytest.raises(ValueError, match="LOTE"):
            Expediente(
                id=uuid4(),
                tipo=ExpedienteTipo.LOTE,
                estado=ExpedienteEstado.BORRADOR,
                id_expediente_padre=None,
                version=1,
                created_at=_now(),
                updated_at=_now(),
            )

    def test_lote_with_parent_is_accepted(self) -> None:
        parent_id = uuid4()
        exp = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.LOTE,
            estado=ExpedienteEstado.BORRADOR,
            id_expediente_padre=parent_id,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert exp.id_expediente_padre == parent_id

    def test_am_has_no_parent_required(self) -> None:
        """AM is the root — no parent reference."""
        exp = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert exp.id_expediente_padre is None

    def test_based_requires_parent(self) -> None:
        """BASED MUST have a parent."""
        with pytest.raises(ValueError, match="BASED"):
            Expediente(
                id=uuid4(),
                tipo=ExpedienteTipo.BASED,
                estado=ExpedienteEstado.BORRADOR,
                id_expediente_padre=None,
                version=1,
                created_at=_now(),
                updated_at=_now(),
            )

    def test_based_with_parent_accepted(self) -> None:
        parent_id = uuid4()
        exp = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.BASED,
            estado=ExpedienteEstado.BORRADOR,
            id_expediente_padre=parent_id,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert exp.id_expediente_padre == parent_id


# ---------------------------------------------------------------------------
# Equality
# ---------------------------------------------------------------------------

class TestExpedienteEquality:
    """Dataclass equality is field-based."""

    def test_same_fields_equal(self) -> None:
        exp_id = uuid4()
        a = Expediente(
            id=exp_id,
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        b = Expediente(
            id=exp_id,
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert a == b

    def test_distinct_id_unequal(self) -> None:
        a = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        b = Expediente(
            id=uuid4(),
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            version=1,
            created_at=_now(),
            updated_at=_now(),
        )
        assert a != b
