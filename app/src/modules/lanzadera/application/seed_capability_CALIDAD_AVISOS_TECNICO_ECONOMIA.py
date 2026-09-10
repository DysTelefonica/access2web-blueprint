# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Capability maps for CALIDAD_AVISOS, TECNICO, ECONOMIA profiles (DA-12, 0003_seed_profiles).

Frozen so accidental mutation raises at import time.
"""

from __future__ import annotations

CALIDAD_AVISOS: dict[str, bool] = {
    "admin": False,
    "calidad": True,
    "calidad_avisos": True,
    "tecnico": False,
    "economia": False,
    "secretaria": False,
}

TECNICO: dict[str, bool] = {
    "admin": False,
    "calidad": False,
    "calidad_avisos": False,
    "tecnico": True,
    "economia": False,
    "secretaria": False,
}

ECONOMIA: dict[str, bool] = {
    "admin": False,
    "calidad": False,
    "calidad_avisos": False,
    "tecnico": False,
    "economia": True,
    "secretaria": False,
}
