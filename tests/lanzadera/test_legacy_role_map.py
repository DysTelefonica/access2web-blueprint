# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #619
"""Tests for legacy_role_map.py — the 7-flag → profile-code matrix (DA-12, H11).

The matrix has 2^7 = 128 possible flag combinations.  We test 16
representative atoms covering:

  - SinAcceso='Sí' exclusivity (DA-12)
  - All-NULL / all-'No' → DEFAULT
  - Single flag active
  - Multiple flags active
  - NULL vs 'No' distinction
"""

from __future__ import annotations

import pytest

from app.src.modules.lanzadera.domain.legacy_role_map import (
    DEFAULT_PROFILE_CODE,
    LEGACY_CODES,
    SIN_ACCESO_PROFILE_CODE,
    resolve_legacy_roles,
    resolve_profile_codes,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def flags(**kwargs: str | None) -> dict[str, str | None]:
    """Build a full 7-flag dict with sensible 'No' defaults."""
    defaults: dict[str, str | None] = {
        "Administrador": "No",
        "Calidad": "No",
        "CalidadAvisos": "No",
        "Técnico": "No",
        "Economía": "No",
        "Secretaría": "No",
        "SinAcceso": "No",
    }
    defaults.update(kwargs)
    return defaults


# ---------------------------------------------------------------------------
# SinAcceso exclusive rule (DA-12)
# ---------------------------------------------------------------------------


class TestSinAccesoExclusive:
    """SinAcceso='Sí' takes exclusive precedence regardless of other flags."""

    def test_sinacceso_solo(self) -> None:
        f = flags(SinAcceso="Sí")
        assert resolve_legacy_roles(f) == (SIN_ACCESO_PROFILE_CODE,)
        assert resolve_profile_codes(f) == [SIN_ACCESO_PROFILE_CODE]

    def test_sinacceso_con_otros_sí(self) -> None:
        f = flags(
            SinAcceso="Sí",
            Administrador="Sí",
            Calidad="Sí",
            Técnico="Sí",
        )
        assert resolve_legacy_roles(f) == (SIN_ACCESO_PROFILE_CODE,)

    def test_sinacceso_con_todos_sí(self) -> None:
        f = flags(
            SinAcceso="Sí",
            Administrador="Sí",
            Calidad="Sí",
            CalidadAvisos="Sí",
            Técnico="Sí",
            Economía="Sí",
            Secretaría="Sí",
        )
        assert resolve_legacy_roles(f) == (SIN_ACCESO_PROFILE_CODE,)


# ---------------------------------------------------------------------------
# DEFAULT: all flags 'No' or NULL
# ---------------------------------------------------------------------------


class TestDefaultCode:
    """All-NULL or all-'No' → DEFAULT."""

    def test_todos_no(self) -> None:
        f = flags()
        assert resolve_legacy_roles(f) == (DEFAULT_PROFILE_CODE,)

    def test_todos_null(self) -> None:
        f = {k: None for k in flags()}
        assert resolve_legacy_roles(f) == (DEFAULT_PROFILE_CODE,)

    def test_solo_sinacceso_no(self) -> None:
        # Only SinAcceso explicitly 'No', others NULL.
        f = {k: None for k in flags()}
        f["SinAcceso"] = "No"
        assert resolve_legacy_roles(f) == (DEFAULT_PROFILE_CODE,)


# ---------------------------------------------------------------------------
# Single flag active
# ---------------------------------------------------------------------------


class TestSingleFlag:
    """One flag = 'Sí', rest = 'No'."""

    @pytest.mark.parametrize(
        "flag_name,expected",
        [
            ("Administrador", "ADMIN"),
            ("Calidad", "CALIDAD"),
            ("CalidadAvisos", "CALIDAD_AVISOS"),
            ("Técnico", "TECNICO"),
            ("Economía", "ECONOMIA"),
            ("Secretaría", "SECRETARIA"),
        ],
    )
    def test_solo_una_bandera_activa(self, flag_name: str, expected: str) -> None:
        f = flags(**{flag_name: "Sí"})
        assert resolve_legacy_roles(f) == (expected,)
        assert resolve_profile_codes(f) == [expected]


# ---------------------------------------------------------------------------
# Multiple flags active (no SinAcceso)
# ---------------------------------------------------------------------------


class TestMultipleFlags:
    """Two or more non-SinAcceso flags = 'Sí'."""

    def test_admin_mas_calidad(self) -> None:
        f = flags(Administrador="Sí", Calidad="Sí")
        codes = resolve_legacy_roles(f)
        assert set(codes) == {"ADMIN", "CALIDAD"}

    def test_tecnico_mas_economia(self) -> None:
        f = flags(Técnico="Sí", Economía="Sí")
        codes = resolve_legacy_roles(f)
        assert set(codes) == {"TECNICO", "ECONOMIA"}

    def test_todos_los_flags_sin_sinacceso(self) -> None:
        f = flags(
            Administrador="Sí",
            Calidad="Sí",
            CalidadAvisos="Sí",
            Técnico="Sí",
            Economía="Sí",
            Secretaría="Sí",
        )
        codes = resolve_legacy_roles(f)
        assert set(codes) == set(LEGACY_CODES) - {SIN_ACCESO_PROFILE_CODE}

    def test_admin_calidad_avisos_tecnico(self) -> None:
        f = flags(
            Administrador="Sí",
            Calidad="Sí",
            CalidadAvisos="Sí",
            Técnico="Sí",
        )
        codes = resolve_legacy_roles(f)
        assert set(codes) == {"ADMIN", "CALIDAD", "CALIDAD_AVISOS", "TECNICO"}


# ---------------------------------------------------------------------------
# NULL vs 'No' distinction
# ---------------------------------------------------------------------------


class TestNullVsNo:
    """NULL and 'No' are treated identically by the resolver."""

    def test_null_es_como_no(self) -> None:
        f_null = {k: None for k in flags()}
        f_no = flags()
        assert resolve_legacy_roles(f_null) == resolve_legacy_roles(f_no) == (DEFAULT_PROFILE_CODE,)

    def test_una_bandera_null(self) -> None:
        f = flags(Calidad=None)
        assert resolve_legacy_roles(f) == (DEFAULT_PROFILE_CODE,)
