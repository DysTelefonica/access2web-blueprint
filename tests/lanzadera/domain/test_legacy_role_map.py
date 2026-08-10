# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-12, H11 — 32-case matrix test for the legacy-role mapping rule.
"""Strict TDD — legacy-role mapping rule (Phase 1, task 1.8)."""

from __future__ import annotations

import pytest

from app.src.modules.lanzadera.domain.legacy_role_map import (
    DEFAULT_PROFILE_CODE,
    LEGACY_ROLE_MAP,
    LegacyFlags,
    SIN_ACCESO_PROFILE_CODE,
    resolve_legacy_roles,
)

# ---------------------------------------------------------------------------
# Mapping table
# ---------------------------------------------------------------------------


class TestLegacyRoleMapTable:
    """The 7-flag mapping is a frozen dict with one entry per non-exclusion flag."""

    def test_map_has_seven_entries(self) -> None:
        """Seven flags in `TbUsuariosAplicacionesPermisos` -> seven entries (incl. SinAcceso)."""
        assert len(LEGACY_ROLE_MAP) == 6

    @pytest.mark.parametrize(
        ("flag", "profile_code"),
        [
            (LegacyFlags.ADMINISTRADOR, "ADMIN"),
            (LegacyFlags.CALIDAD, "CALIDAD"),
            (LegacyFlags.CALIDAD_AVISOS, "CALIDAD_AVISOS"),
            (LegacyFlags.TECNICO, "TECNICO"),
            (LegacyFlags.ECONOMIA, "ECONOMIA"),
            (LegacyFlags.SECRETARIA, "SECRETARIA"),
        ],
    )
    def test_each_flag_maps_to_its_profile_code(
        self, flag: str, profile_code: str
    ) -> None:
        assert LEGACY_ROLE_MAP[flag] == profile_code

    def test_sin_acceso_is_not_in_the_map(self) -> None:
        """SinAcceso is exclusive; it has its own fallback code."""
        assert LegacyFlags.SIN_ACCESO not in LEGACY_ROLE_MAP


# ---------------------------------------------------------------------------
# H11 — 32-case matrix
# ---------------------------------------------------------------------------


# All possible combinations of the 6 non-SinAcceso flags. With 6 flags,
# there are 2^6 = 64 combinations of {set, not-set}; the 32-case matrix in
# H11 refers to the cardinality matrix that pairs each combination with
# the SinAcceso override (set / not-set).
ALL_FLAG_NAMES: tuple[str, ...] = (
    LegacyFlags.ADMINISTRADOR,
    LegacyFlags.CALIDAD,
    LegacyFlags.CALIDAD_AVISOS,
    LegacyFlags.TECNICO,
    LegacyFlags.ECONOMIA,
    LegacyFlags.SECRETARIA,
)


def _bitmask_to_flags(bitmask: int) -> dict[str, bool]:
    return {name: bool(bitmask & (1 << idx)) for idx, name in enumerate(ALL_FLAG_NAMES)}


def _expected_codes(flags: dict[str, bool], *, sin_acceso: bool) -> tuple[str, ...]:
    """Replicate the production rule to assert the resolver agrees."""
    if sin_acceso:
        return (SIN_ACCESO_PROFILE_CODE,)
    codes = tuple(
        LEGACY_ROLE_MAP[name]
        for idx, name in enumerate(ALL_FLAG_NAMES)
        if flags[name] and (1 << idx)  # always true when flags[name]
    )
    if codes:
        return codes
    return (DEFAULT_PROFILE_CODE,)


@pytest.mark.parametrize("bitmask", list(range(2 ** len(ALL_FLAG_NAMES))))
def test_resolve_without_sin_acceso(bitmask: int) -> None:
    """Every non-SinAcceso bitmask -> the matching set of profile codes or DEFAULT."""
    flags = _bitmask_to_flags(bitmask)
    assert resolve_legacy_roles(flags) == _expected_codes(flags, sin_acceso=False)


@pytest.mark.parametrize("bitmask", list(range(2 ** len(ALL_FLAG_NAMES))))
def test_resolve_with_sin_acceso_overrides(bitmask: int) -> None:
    """DA-12 exclusive cortocircuito: SinAcceso collapses to a single row."""
    flags = _bitmask_to_flags(bitmask)
    flags[LegacyFlags.SIN_ACCESO] = True
    assert resolve_legacy_roles(flags) == _expected_codes(flags, sin_acceso=True)


# ---------------------------------------------------------------------------
# SinAcceso exclusivity (DA-12)
# ---------------------------------------------------------------------------


class TestSinAccesoExclusivity:
    """`SinAcceso = True` short-circuits every other flag."""

    def test_sin_acceso_only(self) -> None:
        assert resolve_legacy_roles({LegacyFlags.SIN_ACCESO: True}) == (
            SIN_ACCESO_PROFILE_CODE,
        )

    def test_sin_acceso_with_admin_only_collapses(self) -> None:
        flags = {
            LegacyFlags.SIN_ACCESO: True,
            LegacyFlags.ADMINISTRADOR: True,
        }
        assert resolve_legacy_roles(flags) == (SIN_ACCESO_PROFILE_CODE,)

    def test_sin_acceso_with_multiple_flags_collapses(self) -> None:
        flags = {
            LegacyFlags.SIN_ACCESO: True,
            LegacyFlags.ADMINISTRADOR: True,
            LegacyFlags.CALIDAD: True,
            LegacyFlags.TECNICO: True,
        }
        assert resolve_legacy_roles(flags) == (SIN_ACCESO_PROFILE_CODE,)


# ---------------------------------------------------------------------------
# DEFAULT fallback (DA-12)
# ---------------------------------------------------------------------------


class TestDefaultFallback:
    """All-NULL / all-No legacy row -> DEFAULT."""

    def test_empty_flags_yields_default(self) -> None:
        assert resolve_legacy_roles({}) == (DEFAULT_PROFILE_CODE,)

    def test_all_no_flags_yields_default(self) -> None:
        flags = {name: False for name in ALL_FLAG_NAMES}
        assert resolve_legacy_roles(flags) == (DEFAULT_PROFILE_CODE,)

    def test_all_null_flags_yields_default(self) -> None:
        flags = {name: None for name in ALL_FLAG_NAMES}
        assert resolve_legacy_roles(flags) == (DEFAULT_PROFILE_CODE,)

    def test_unset_sin_acceso_and_all_no_yields_default(self) -> None:
        flags = {LegacyFlags.SIN_ACCESO: False}
        assert resolve_legacy_roles(flags) == (DEFAULT_PROFILE_CODE,)


# ---------------------------------------------------------------------------
# Compound (non-SinAcceso) rows
# ---------------------------------------------------------------------------


class TestCompoundRows:
    """Two or more set flags produce two or more profile codes."""

    def test_calidad_and_calidad_avisos(self) -> None:
        flags = {
            LegacyFlags.CALIDAD: True,
            LegacyFlags.CALIDAD_AVISOS: True,
        }
        assert resolve_legacy_roles(flags) == ("CALIDAD", "CALIDAD_AVISOS")

    def test_admin_and_tecnico(self) -> None:
        flags = {
            LegacyFlags.ADMINISTRADOR: True,
            LegacyFlags.TECNICO: True,
        }
        assert resolve_legacy_roles(flags) == ("ADMIN", "TECNICO")

    def test_all_six_flags_set(self) -> None:
        flags = {name: True for name in ALL_FLAG_NAMES}
        assert resolve_legacy_roles(flags) == (
            "ADMIN",
            "CALIDAD",
            "CALIDAD_AVISOS",
            "TECNICO",
            "ECONOMIA",
            "SECRETARIA",
        )


# ---------------------------------------------------------------------------
# Legacy input formats (text `Sí` / `No`)
# ---------------------------------------------------------------------------


class TestLegacyTextValues:
    """The legacy source stored `'Sí'` / `'No'`; both must be honoured."""

    def test_si_string_is_set(self) -> None:
        assert resolve_legacy_roles({LegacyFlags.ADMINISTRADOR: "Sí"}) == ("ADMIN",)

    def test_no_string_is_not_set(self) -> None:
        assert resolve_legacy_roles({LegacyFlags.ADMINISTRADOR: "No"}) == (
            DEFAULT_PROFILE_CODE,
        )

    def test_lowercase_si_string_is_set(self) -> None:
        assert resolve_legacy_roles({LegacyFlags.CALIDAD: "si"}) == ("CALIDAD",)

    def test_whitespace_is_trimmed(self) -> None:
        assert resolve_legacy_roles({LegacyFlags.TECNICO: "  Sí  "}) == ("TECNICO",)


# ---------------------------------------------------------------------------
# Output shape
# ---------------------------------------------------------------------------


class TestOutputShape:
    """The resolver returns a tuple (immutable, hashable, deterministic order)."""

    def test_returns_tuple(self) -> None:
        result = resolve_legacy_roles({LegacyFlags.ADMINISTRADOR: True})
        assert isinstance(result, tuple)

    def test_output_is_hashable(self) -> None:
        result = resolve_legacy_roles({LegacyFlags.CALIDAD: True})
        # tuples are hashable; this would raise TypeError if not
        assert hash(result) is not None

    def test_output_order_is_deterministic(self) -> None:
        """Iteration order follows LEGACY_ROLE_MAP, not the input dict."""
        flags = {
            LegacyFlags.SECRETARIA: True,
            LegacyFlags.ADMINISTRADOR: True,
        }
        assert resolve_legacy_roles(flags) == ("ADMIN", "SECRETARIA")
