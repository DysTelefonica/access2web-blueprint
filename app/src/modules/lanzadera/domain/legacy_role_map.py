# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-12, H11, D110 — pure domain module; only stdlib imports allowed.
"""Legacy-role mapping (lanzadera-mvp/assignments).

DA-12: the 7-flag legacy `TbUsuariosAplicacionesPermisos` table collapses
to profile codes per `(user, app)` triple. The mapping is declared as a
frozen dict-of-sets so the rule is auditable from a single read; the
runtime function applies the SinAcceso exclusivity cortocircuito and the
all-NULL `DEFAULT` fallback.

H11 (32-case matrix): the function is exercised against every
combination of the 6 non-SinAcceso flags plus the SinAcceso override.
D110 forbids inferring capabilities from dynamic fields; the canonical
shape arrives in PR 3a (migration 0003).
"""

from __future__ import annotations

from enum import StrEnum

__all__ = [
    "LegacyFlags",
    "LEGACY_ROLE_MAP",
    "DEFAULT_PROFILE_CODE",
    "SIN_ACCESO_PROFILE_CODE",
    "resolve_legacy_roles",
]


# Wire values MUST match the legacy `TbUsuariosAplicacionesPermisos` columns.
class LegacyFlags(StrEnum):
    """Boolean flag columns from the legacy permission table."""

    ADMINISTRADOR = "Administrador"
    CALIDAD = "Calidad"
    CALIDAD_AVISOS = "CalidadAvisos"
    TECNICO = "Técnico"
    ECONOMIA = "Economía"
    SECRETARIA = "Secretaría"
    SIN_ACCESO = "SinAcceso"


# Mapping from each non-exclusion flag to the modern profile code.
LEGACY_ROLE_MAP: dict[str, str] = {
    LegacyFlags.ADMINISTRADOR: "ADMIN",
    LegacyFlags.CALIDAD: "CALIDAD",
    LegacyFlags.CALIDAD_AVISOS: "CALIDAD_AVISOS",
    LegacyFlags.TECNICO: "TECNICO",
    LegacyFlags.ECONOMIA: "ECONOMIA",
    LegacyFlags.SECRETARIA: "SECRETARIA",
}

# Fallback codes.
DEFAULT_PROFILE_CODE: str = "DEFAULT"
SIN_ACCESO_PROFILE_CODE: str = "SIN_ACCESO"


def _flag_is_set(flags: dict[str, bool | None], name: str) -> bool:
    """A flag is set when its legacy value is truthy (`'Sí'`, `True`, ...).

    The legacy source stored `'Sí'` / `'No'` text; this helper coerces both
    the boolean and the text form so adapters can pass either without
    branching.
    """
    value = flags.get(name)
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"sí", "si", "yes", "true", "1"}
    return bool(value)


def resolve_legacy_roles(flags: dict[str, bool | None]) -> tuple[str, ...]:
    """Map one legacy `(user, app)` permission row to its modern profile codes.

    The mapping follows the assignment spec (H11):

    1. If `SinAcceso` is set, the row collapses to a single `SIN_ACCESO`
       row — every other flag is ignored (DA-12 exclusive cortocircuito).
    2. Otherwise, every set non-exclusion flag contributes its mapped
       profile code. Multiple set flags produce multiple rows.
    3. If no flag is set (all-No / all-NULL), the row collapses to the
       `DEFAULT` profile code.

    Returns a tuple (not a list) so the output is hashable and a
    downstream migration can iterate it deterministically.
    """
    if _flag_is_set(flags, LegacyFlags.SIN_ACCESO):
        return (SIN_ACCESO_PROFILE_CODE,)

    active = tuple(LEGACY_ROLE_MAP[flag] for flag in LEGACY_ROLE_MAP if _flag_is_set(flags, flag))
    if active:
        return active

    return (DEFAULT_PROFILE_CODE,)
