# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Capability maps for DEFAULT, ADMIN, CALIDAD profiles (DA-12, 0003_seed_profiles).

Each dict maps 6 legacy-role booleans.
Frozen so accidental mutation raises at import time.

The JSONB contract accepts any ``str | int | bool`` values.
"""

from __future__ import annotations

DEFAULT: dict[str, bool] = {
    "admin": False,
    "calidad": False,
    "calidad_avisos": False,
    "tecnico": False,
    "economia": False,
    "secretaria": False,
}

ADMIN: dict[str, bool] = {
    "admin": True,
    "calidad": True,
    "calidad_avisos": True,
    "tecnico": True,
    "economia": True,
    "secretaria": True,
}

CALIDAD: dict[str, bool] = {
    "admin": False,
    "calidad": True,
    "calidad_avisos": False,
    "tecnico": False,
    "economia": False,
    "secretaria": False,
}
