# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Static profile-code definitions (DA-12, 0003_seed_profiles).

Profile codes map to legacy role booleans via ``_CAPABILITIES_BY_CODE``
(the capability data lives in the ``seed_capability_*.py`` group modules).

Codes (from ``legacy_role_map.LEGACY_ROLE_MAP``):
    DEFAULT        — fallback when no explicit profile matches
    ADMIN          — maps to legacy EsAdministrador
    CALIDAD        — maps to legacy EsCalidad
    CALIDAD_AVISOS — maps to legacy EsCalidadAvisos
    TECNICO        — maps to legacy EsTecnico
    ECONOMIA       — maps to legacy EsEconomía
    SECRETARIA     — maps to legacy EsSecretaría
    SIN_ACCESO     — exclusive: blocks all access (DA-12 cortocircuito)
"""

from __future__ import annotations

#: Profile codes and their display names.
#: Frozen so accidental mutation raises at import time.
PROFILE_CODES: tuple[dict[str, str], ...] = (
    {"code": "DEFAULT", "name": "Usuario por defecto"},
    {"code": "ADMIN", "name": "Administrador"},
    {"code": "CALIDAD", "name": "Calidad"},
    {"code": "CALIDAD_AVISOS", "name": "Calidad + Avisos"},
    {"code": "TECNICO", "name": "Técnico"},
    {"code": "ECONOMIA", "name": "Economía"},
    {"code": "SECRETARIA", "name": "Secretaría"},
    {"code": "SIN_ACCESO", "name": "Sin acceso"},
)
