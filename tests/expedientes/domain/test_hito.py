"""Strict TDD — `Hito` child entity (F02, issue #223).

RED: tests first, then Hito entity.
"""

from __future__ import annotations

from datetime import UTC, datetime, date
from uuid import uuid4

import pytest

from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado


def _d(year: int, month: int, day: int) -> date:
    return date(year, month, day)


def _dt(year: int, month: int, day: int) -> datetime:
    return datetime(year, month, day, 12, 0, 0, tzinfo=UTC)


# ---------------------------------------------------------------------------
# HitoEstado StrEnum
# ---------------------------------------------------------------------------


class TestHitoEstadoStrEnum:
    """Must be a StrEnum matching the Postgres enum literal values."""

    def test_members_are_str(self) -> None:
        from app.src.modules.expedientes.domain.hito import HitoEstado

        for member in HitoEstado:
            assert isinstance(member, str)

    def test_pendiente_value(self) -> None:
        from app.src.modules.expedientes.domain.hito import HitoEstado

        assert HitoEstado.PENDIENTE.value == "PENDIENTE"

    def test_cumplido_value(self) -> None:
        from app.src.modules.expedientes.domain.hito import HitoEstado

        assert HitoEstado.CUMPLIDO.value == "CUMPLIDO"

    def test_vencido_value(self) -> None:
        from app.src.modules.expedientes.domain.hito import HitoEstado

        assert HitoEstado.VENCIDO.value == "VENCIDO"


# ---------------------------------------------------------------------------
# Hito construction — happy path
# ---------------------------------------------------------------------------


class TestHitoConstruction:
    """Minimum viable `Hito` carries identity, dates and estado."""

    def test_carries_identity_attributes(self) -> None:
        from app.src.modules.expedientes.domain.hito import Hito

        id_expediente = uuid4()
        hito = Hito(
            id=uuid4(),
            id_expediente=id_expediente,
            fecha_hito=_d(2026, 10, 1),
            garantia_fecha_fin=_d(2027, 10, 1),
            estado=ExpedienteEstado.BORRADOR,
            created_at=_dt(2026, 9, 1),
            updated_at=_dt(2026, 9, 1),
        )
        assert hito.id_expediente == id_expediente
        assert hito.fecha_hito == _d(2026, 10, 1)
        assert hito.garantia_fecha_fin == _d(2027, 10, 1)


# ---------------------------------------------------------------------------
# Hito invariants
# ---------------------------------------------------------------------------


class TestHitoInvariants:
    """Domain invariants enforced at construction time (CAP-008)."""

    def test_fecha_hito_required(self) -> None:
        """fecha_hito is mandatory — cannot be None."""
        from app.src.modules.expedientes.domain.hito import Hito

        with pytest.raises(ValueError, match="fecha_hito"):
            Hito(
                id=uuid4(),
                id_expediente=uuid4(),
                fecha_hito=None,  # type: ignore
                garantia_fecha_fin=None,
                estado=ExpedienteEstado.BORRADOR,
                created_at=_dt(2026, 9, 1),
                updated_at=_dt(2026, 9, 1),
            )

    def test_fecha_hito_must_be_date_not_datetime(self) -> None:
        """fecha_hito is a date (no time component)."""
        from app.src.modules.expedientes.domain.hito import Hito

        with pytest.raises(ValueError, match="date"):
            Hito(
                id=uuid4(),
                id_expediente=uuid4(),
                fecha_hito=_dt(2026, 10, 1),  # datetime, not date
                garantia_fecha_fin=None,
                estado=ExpedienteEstado.BORRADOR,
                created_at=_dt(2026, 9, 1),
                updated_at=_dt(2026, 9, 1),
            )

    def test_garantia_fecha_fin_after_fecha_hito(self) -> None:
        """garantia_fecha_fin must be strictly after fecha_hito (CAP-008)."""
        from app.src.modules.expedientes.domain.hito import Hito

        with pytest.raises(ValueError, match="garantia"):
            Hito(
                id=uuid4(),
                id_expediente=uuid4(),
                fecha_hito=_d(2026, 10, 1),
                garantia_fecha_fin=_d(2026, 10, 1),  # same day — not strictly after
                estado=ExpedienteEstado.BORRADOR,
                created_at=_dt(2026, 9, 1),
                updated_at=_dt(2026, 9, 1),
            )

    def test_garantia_fecha_fin_before_fecha_hito_rejected(self) -> None:
        """garantia_fecha_fin must be strictly after fecha_hito."""
        from app.src.modules.expedientes.domain.hito import Hito

        with pytest.raises(ValueError, match="garantia"):
            Hito(
                id=uuid4(),
                id_expediente=uuid4(),
                fecha_hito=_d(2026, 12, 1),
                garantia_fecha_fin=_d(2026, 10, 1),  # before fecha_hito
                estado=ExpedienteEstado.BORRADOR,
                created_at=_dt(2026, 9, 1),
                updated_at=_dt(2026, 9, 1),
            )

    def test_garantia_fecha_fin_can_be_none(self) -> None:
        """garantia_fecha_fin is optional (no warranty milestone)."""
        from app.src.modules.expedientes.domain.hito import Hito

        hito = Hito(
            id=uuid4(),
            id_expediente=uuid4(),
            fecha_hito=_d(2026, 10, 1),
            garantia_fecha_fin=None,
            estado=ExpedienteEstado.BORRADOR,
            created_at=_dt(2026, 9, 1),
            updated_at=_dt(2026, 9, 1),
        )
        assert hito.garantia_fecha_fin is None

    def test_fecha_hito_valid_date_accepted(self) -> None:
        """A valid date for fecha_hito is accepted."""
        from app.src.modules.expedientes.domain.hito import Hito

        hito = Hito(
            id=uuid4(),
            id_expediente=uuid4(),
            fecha_hito=_d(2026, 10, 1),
            garantia_fecha_fin=_d(2027, 10, 2),
            estado=ExpedienteEstado.BORRADOR,
            created_at=_dt(2026, 9, 1),
            updated_at=_dt(2026, 9, 1),
        )
        assert hito.fecha_hito == _d(2026, 10, 1)
