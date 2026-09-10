# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W-TEST
"""Static app catalogue for the 8 known legacy apps (DA-12, 0003_seed_profiles).

IDs are confirmed in ``docs/06-autorizacion-legacy-matriz.md``
(Verified-runtime-schema/aggregate).

Each entry maps to a row in ``lanzadera/apps`` (app_id column).
The ``short_code`` is the canonical legacy identifier used in the
capability-gate lookup (DA-12 / ``0003_seed_profiles`` spec).
"""

from __future__ import annotations

#: Eight apps from the legacy ``TbAplicaciones`` baseline.
#: Frozen so accidental mutation raises at import time.
APP_CATALOGUE: tuple[dict[str, int | str], ...] = (
    {"id": 5, "name": "Gestion_Riesgos", "short_code": "RIESGOS"},
    {"id": 6, "name": "Brass", "short_code": "BRASS"},
    {"id": 8, "name": "No_Conformidades", "short_code": "NOCONF"},
    {"id": 12, "name": "Lanzadera", "short_code": "LANZ"},
    {"id": 17, "name": "HPS", "short_code": "HPS"},
    {"id": 19, "name": "Expedientes", "short_code": "EXP"},
    {"id": 22, "name": "HPS_Solicitudes", "short_code": "HPSSOL"},
    {"id": 23, "name": "Condor", "short_code": "CONDOR"},
)
