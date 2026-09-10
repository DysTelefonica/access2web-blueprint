# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Capability maps for SECRETARIA, SIN_ACCESO profiles (DA-12, 0003_seed_profiles).

Frozen so accidental mutation raises at import time.
"""

from __future__ import annotations

SECRETARIA: dict[str, bool] = {
    "admin": False,
    "calidad": False,
    "calidad_avisos": False,
    "tecnico": False,
    "economia": False,
    "secretaria": True,
}

SIN_ACCESO: dict[str, bool] = {
    "admin": False,
    "calidad": False,
    "calidad_avisos": False,
    "tecnico": False,
    "economia": False,
    "secretaria": False,
}
