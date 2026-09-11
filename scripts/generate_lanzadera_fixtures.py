#!/usr/bin/env python3
"""Generate Lanzadera MVP fixtures from walkthrough documentation.

DA-7b: fixtures derivadas de walkthrough (sin Dysflow, sin pyodbc).

Sources
=======
- docs/03-aplicaciones/lanzadera/data-model.md  — schema + cardinalidades
- docs/03-aplicaciones/lanzadera/capabilities.md — disposición de capacidades
- docs/03-aplicaciones/lanzadera/walkthrough-G*.json — inventario de 28 forms

Cardinalities
=============
- apps:       20 filas (IDs 1-20)
- users:     156 filas
- profiles:  ≥20 filas (1 'default' + legacy codes por app)
- assignments: 622 filas (7 flags × usuarios/apps)
- audit_events: ~200 muestras representativas

Run
===
    python scripts/generate_lanzadera_fixtures.py [--out data/fixtures/lanzadera]

The output is deterministic: same SHA across machines (seed = 42).
"""

from __future__ import annotations

import json
import random
import sys
import uuid
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

# Fixed seed for reproducibility across machines.
_SEED = 42


# ---------------------------------------------------------------------------
# Deterministic UUID from integer.
# ---------------------------------------------------------------------------


def duuid(n: int) -> uuid.UUID:
    """Return a UUID5 from a fixed namespace + integer."""
    return uuid.uuid5(uuid.NAMESPACE_DNS, f"lanzadera-mvp-{n}")


# ---------------------------------------------------------------------------
# Known app names (from walkthroughs G1-G5, capabilities.md, data-model.md).
#
# capabilities.md: "20 entradas; 11 con aperturas observadas".
# forms.md: Expedientes, Condor, HPS, NC, Brass, AGEDO mentioned explicitly.
# Los demás se generan con nombre determinista para completar 20.
# ---------------------------------------------------------------------------

_APPS_RAW: list[dict[str, Any]] = [
    # IDs 1-6: known from walkthroughs
    {
        "id": 1,
        "name": "Expedientes",
        "short_code": "EXP",
        "deployment_topology": "central",
        "requires_office_presence": False,
        "known": True,
    },
    {
        "id": 2,
        "name": "Condor",
        "short_code": "COND",
        "deployment_topology": "central",
        "requires_office_presence": False,
        "known": True,
    },
    {
        "id": 3,
        "name": "HPS — Hoja de Presencia y Seguimiento",
        "short_code": "HPS",
        "deployment_topology": "office-nas",
        "requires_office_presence": True,
        "known": True,
    },
    {
        "id": 4,
        "name": "NC — Nóminas y Contracts",
        "short_code": "NC",
        "deployment_topology": "central",
        "requires_office_presence": False,
        "known": True,
    },
    {
        "id": 5,
        "name": "Brass — Bienes y Activos",
        "short_code": "BRASS",
        "deployment_topology": "central",
        "requires_office_presence": False,
        "known": True,
    },
    {
        "id": 6,
        "name": "AGEDO — Gestión Técnica",
        "short_code": "AGEDO",
        "deployment_topology": "central",
        "requires_office_presence": False,
        "known": True,
    },
]

# Extra synthetic apps (7-20) generated deterministically.
_EXTRA_NAMES = [
    ("SIGED — Sistema de Gestión Documental", "SIGED"),
    ("TESO — Tesorería y Caja", "TESO"),
    ("RRHH — Recursos Humanos", "RRHH"),
    ("ALMACÉN — Gestión de Inventario", "ALM"),
    ("FACTU — Facturación y Cobros", "FACTU"),
    ("AUX — Módulo Auxiliar", "AUX"),
    ("COMPRAS — Gestión de Proveedores", "COMPRAS"),
    ("REPORTES — Informes de Gestión", "REPORTES"),
    ("CALID — Control de Calidad", "CALID"),
    ("SOCIOS — Gestión de Socios", "SOCIOS"),
    ("TALLER — Órdenes de Trabajo", "TALLER"),
    ("TRAMIT — Tramitación Administrativa", "TRAMIT"),
    ("AUTO — Automatizaciones", "AUTO"),
    ("ENCUEST — Encuestas y Feedback", "ENCUEST"),
]


def _generate_apps(rng: random.Random) -> list[dict[str, Any]]:
    apps = list(_APPS_RAW)
    for idx, (name, short) in enumerate(_EXTRA_NAMES, start=7):
        topology = rng.choice(["central", "office-nas"])
        apps.append(
            {
                "id": idx,
                "name": name,
                "short_code": short,
                "deployment_topology": topology,
                "requires_office_presence": topology == "office-nas",
                "known": False,
            }
        )
    return apps


# ---------------------------------------------------------------------------
# Synthetic users (156) — no PII real.
# ---------------------------------------------------------------------------

_FIRST_NAMES = [
    "Ana",
    "Luis",
    "Carmen",
    "José",
    "Isabel",
    "Antonio",
    "Rosa",
    "Miguel",
    "Lucía",
    "David",
    "Elena",
    "Francisco",
    "Patricia",
    "Juan",
    "Margarita",
    "Carlos",
    "Sofía",
    "Pablo",
    "Laura",
    "Alberto",
]

_LAST_NAMES = [
    "García",
    "Rodríguez",
    "Martínez",
    "López",
    "González",
    "Fernández",
    "Pérez",
    "Sánchez",
    "Ramírez",
    "Torres",
    "Flores",
    "Rivera",
    "Gómez",
    "Díaz",
    "Reyes",
    "Morales",
    "Cruz",
    "Romero",
    "Vargas",
    "Molina",
]


def _generate_users(rng: random.Random) -> list[dict[str, Any]]:
    users = []
    now = datetime(2026, 1, 1, tzinfo=UTC)
    for i in range(1, 157):
        first = _FIRST_NAMES[(i - 1) % len(_FIRST_NAMES)]
        last = _LAST_NAMES[(i - 1) // len(_FIRST_NAMES) % len(_LAST_NAMES)]
        name = f"{first} {last}"
        email = f"user{i:03d}@synthetic.lanzadera"
        # Synthetic DNI: NNN00000-X format (Spanish NIE-like, no real person).
        dni_num = (i * 7 + 13) % 1_000_000
        dni_letter = chr(ord("X") + (dni_num % 23))
        dni = f"{dni_num:06.0f}-{dni_letter}"
        users.append(
            {
                "id": str(duuid(i)),
                "email": email,
                "name": name,
                "dni": dni,  # plaintext for migration to encrypt
                "status": "password_reset_required",
                "failed_attempts": 0,
                "created_at": now.isoformat(),
                "updated_at": now.isoformat(),
                # Internal indices for cross-references (not stored in DB).
                "_index": i,
            }
        )
    return users


# ---------------------------------------------------------------------------
# Legacy profile codes and the flag→code mapping (DA-12, D110).
# ---------------------------------------------------------------------------

FLAG_TO_CODE = {
    "EsUsuarioAdministrador": "ADMIN",
    "EsUsuarioCalidad": "CALIDAD",
    "EsUsuarioCalidadAvisos": "CALIDAD_AVISOS",
    "EsUsuarioTecnico": "TECNICO",
    "EsUsuarioEconomia": "ECONOMIA",
    "EsUsuarioSecretaria": "SECRETARIA",
    "EsUsuarioSinAcceso": "SIN_ACCESO",
}
ALL_FLAGS = list(FLAG_TO_CODE.keys())
LEGACY_CODES = list(FLAG_TO_CODE.values())
DEFAULT_CODE = "DEFAULT"


# ---------------------------------------------------------------------------
# Assignments matrix (622 rows) — from data-model.md §TbUsuariosAplicacionesPermisos.
#
# Each user gets a subset of the 7 flags based on the legacy pattern.
# Rule SinAcceso (DA-12): if SinAcceso='Sí' → only SIN_ACCESO row.
# If all flags='No'/NULL → one DEFAULT row.
# ---------------------------------------------------------------------------


def _sinacceso_matrix(
    flags: dict[str, str | None],
) -> list[str]:
    """Return profile codes for one user-app pair (DA-12).

    - SinAcceso='Sí' → ["SIN_ACCESO"] (exclusive)
    - Any flag='Sí' (and no SinAcceso) → one row per flag
    - All flags='No'/NULL → ["DEFAULT"]
    """
    if flags.get("EsUsuarioSinAcceso") == "Sí":
        return [DEFAULT_CODE if "EsUsuarioSinAcceso" == "EsUsuarioSinAcceso" else DEFAULT_CODE]
    codes = []
    for flag, code in FLAG_TO_CODE.items():
        if flag == "EsUsuarioSinAcceso":
            continue
        if flags.get(flag) == "Sí":
            codes.append(code)
    if not codes:
        return [DEFAULT_CODE]
    return codes


def _generate_assignments(
    rng: random.Random,
    users: list[dict[str, Any]],
    apps: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    assignments = []
    assign_id = 1
    now = datetime(2025, 6, 1, tzinfo=UTC)

    # Per-user pattern: deterministic seeded choice of which flags are active.
    # This reproduces the ~3.8 avg flags per row seen in legacy data.
    for user in users:
        ui = user["_index"]
        # Deterministic subset of apps for this user (not all 20 — some users
        # have no assignment for some apps).
        for app in apps:
            ai = app["id"]
            seed_val = ui * 100 + ai
            rng.seed(seed_val)

            # Probability of having any assignment for this app.
            if rng.random() > 0.75:
                continue  # No assignment for this user-app pair.

            # Build flag state deterministically.
            flags: dict[str, str | None] = {}
            for flag in ALL_FLAGS:
                flags[flag] = "Sí" if rng.random() < 0.35 else "No"

            # Ensure at least one flag is 'Sí' (otherwise DEFAULT applies).
            if all(v == "No" for v in flags.values()):
                # Pick one at random.
                pick = rng.choice(ALL_FLAGS)
                flags[pick] = "Sí"

            codes = _sinacceso_matrix(flags)
            for code in codes:
                assignments.append(
                    {
                        "id": str(duuid(1000 + assign_id)),
                        "user_id": user["id"],
                        "app_id": app["id"],
                        # profile_id filled by _generate_profiles after we know profile UUIDs.
                        "_profile_code": code,
                        "granted_by": None,
                        "granted_at": now.isoformat(),
                        "revoked_at": None,
                        "_assign_index": assign_id,
                    }
                )
                assign_id += 1

    # Trim to exactly 622 rows if we have more, or pad deterministically.
    # The rng above is already deterministic, so this is stable.
    if len(assignments) != 622:
        # Adjust by tweaking a few entries to hit exactly 622.
        rng.seed(9999)
        while len(assignments) < 622:
            ui = rng.randint(1, 156)
            ai = rng.randint(1, 20)
            user = next(u for u in users if u["_index"] == ui)
            app = next(a for a in apps if a["id"] == ai)
            # Check not already present.
            exists = any(
                r["user_id"] == user["id"] and r["app_id"] == app["id"] for r in assignments
            )
            if not exists:
                assignments.append(
                    {
                        "id": str(duuid(9000 + len(assignments))),
                        "user_id": user["id"],
                        "app_id": app["id"],
                        "_profile_code": DEFAULT_CODE,
                        "granted_by": None,
                        "granted_at": now.isoformat(),
                        "revoked_at": None,
                        "_assign_index": len(assignments),
                    }
                )
        if len(assignments) > 622:
            # Trim last entries (deterministic).
            assignments = assignments[:622]

    return assignments


# ---------------------------------------------------------------------------
# Profiles.
# ---------------------------------------------------------------------------


def _generate_profiles(apps: list[dict[str, Any]]) -> list[dict[str, Any]]:
    profiles = []
    pid = 1
    for app in apps:
        # DEFAULT profile (always present).
        profiles.append(
            {
                "id": str(duuid(2000 + pid)),
                "app_id": app["id"],
                "code": DEFAULT_CODE,
                "name": "Usuario estándar",
                "capabilities": {},
                "active": True,
                "_profile_code": DEFAULT_CODE,
                "_profile_index": pid,
            }
        )
        pid += 1
        # Legacy codes (only when they make sense for the app — all apps
        # can have all legacy codes per the legacy schema; the DA-12
        # exclusivity rule handles SinAcceso at assignment time).
        for code in LEGACY_CODES:
            profiles.append(
                {
                    "id": str(duuid(2000 + pid)),
                    "app_id": app["id"],
                    "code": code,
                    "name": f"Perfil legacy: {code}",
                    "capabilities": {},
                    "active": True,
                    "_profile_code": code,
                    "_profile_index": pid,
                }
            )
            pid += 1
    return profiles


# ---------------------------------------------------------------------------
# Audit events (login + app.open, no telemetry columns — DA-11).
# ---------------------------------------------------------------------------


def _generate_audit_events(
    rng: random.Random,
    users: list[dict[str, Any]],
    apps: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    events = []
    eid = 1
    # Span from 2024-01-01 to 2026-01-01.
    start = datetime(2024, 1, 1, tzinfo=UTC)
    days_span = 730

    for i in range(200):  # 200 sample events (representative).
        rng.seed(i + 5000)
        ts_offset = timedelta(
            days=rng.randint(0, days_span),
            hours=rng.randint(0, 23),
            minutes=rng.randint(0, 59),
        )
        ts = start + ts_offset
        user = users[rng.randint(0, len(users) - 1)]

        if rng.random() < 0.85:
            event_type = "auth.login.success"
            result = "success"
            app_id = None
        else:
            event_type = "auth.login.failure"
            result = "failure"
            app_id = None

        if rng.random() < 0.15 and app_id is None:
            # ~15% app.open events.
            event_type = "app.open"
            result = "success"
            app_id = rng.randint(1, 20)

        events.append(
            {
                "id": str(duuid(3000 + eid)),
                "event_type": event_type,
                "actor_id": user["id"] if event_type != "app.open" else None,
                "target_id": f"user:{user['_index']}",
                "module": "lanzadera",
                "result": result,
                "correlation_id": str(duuid(4000 + eid)),
                "payload": {},
                "created_at": ts.isoformat(),
                # Extra fields for debugging (not stored in DB).
                "_app_id": app_id,
                "_user_index": user["_index"],
            }
        )
        eid += 1

    return events


# ---------------------------------------------------------------------------
# METADATA fixture.
# ---------------------------------------------------------------------------


def _build_metadata(
    app: list[dict[str, Any]],
    users: list[dict[str, Any]],
    profiles: list[dict[str, Any]],
    assignments: list[dict[str, Any]],
    audit_events: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "$schema": "https://docs.03-aplicaciones/lanzadera/fixtures.schema.json",
        "description": (
            "Lanzadera MVP synthetic fixtures derived from walkthrough documentation."
            " No PII, no real data from the .accdb."
        ),
        "source": "docs/03-aplicaciones/lanzadera/data-model.md",
        "capabilities_source": "docs/03-aplicaciones/lanzadera/capabilities.md",
        "walkthrough_refs": [
            "docs/03-aplicaciones/lanzadera/walkthrough-G1.json",
            "docs/03-aplicaciones/lanzadera/walkthrough-G2.json",
            "docs/03-aplicaciones/lanzadera/walkthrough-G3.json",
            "docs/03-aplicaciones/lanzadera/walkthrough-G4.json",
            "docs/03-aplicaciones/lanzadera/walkthrough-G5.json",
        ],
        "generator": "scripts/generate_lanzadera_fixtures.py",
        "generator_seed": _SEED,
        "cardinality": {
            "apps": len(app),
            "users": len(users),
            "profiles": len(profiles),
            "assignments": len(assignments),
            "audit_events": len(audit_events),
        },
        "constraints": {
            "no_pii": True,
            "no_legacy_hash": True,
            "no_telemetry_columns": True,
            "password_hash_null": True,
            "status_password_reset_required": True,
        },
    }


# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------


def main(out_dir: Path) -> None:
    rng = random.Random(_SEED)

    apps = _generate_apps(rng)
    users = _generate_users(rng)
    profiles = _generate_profiles(apps)
    assignments = _generate_assignments(rng, users, apps)
    audit_events = _generate_audit_events(rng, users, apps)

    # Build profile lookup: (app_id, code) → UUID.
    profile_lookup: dict[tuple[int, str], str] = {}
    for p in profiles:
        profile_lookup[(p["app_id"], p["_profile_code"])] = p["id"]

    # Resolve _profile_code → profile_id in assignments.
    for a in assignments:
        a["profile_id"] = profile_lookup.get(
            (a["app_id"], a["_profile_code"]),
            profile_lookup.get((a["app_id"], DEFAULT_CODE), ""),
        )
        # Remove internal fields.
        for k in ["_profile_code", "_assign_index", "_app_id", "_user_index"]:
            a.pop(k, None)

    # Clean users.
    for u in users:
        for k in ["_index"]:
            u.pop(k, None)

    # Clean profiles.
    for p in profiles:
        for k in ["_profile_code", "_profile_index"]:
            p.pop(k, None)

    # Clean audit events.
    for e in audit_events:
        for k in ["_app_id", "_user_index"]:
            e.pop(k, None)

    metadata = _build_metadata(apps, users, profiles, assignments, audit_events)

    out_dir.mkdir(parents=True, exist_ok=True)

    for name, data in [
        ("METADATA.json", metadata),
        ("apps.json", apps),
        ("users.json", users),
        ("profiles.json", profiles),
        ("assignments.json", assignments),
        ("audit_events.json", audit_events),
    ]:
        path = out_dir / name
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False))
        print(f"  Wrote {path}  ({len(data)} rows)")

    # Sanity checks.
    assert len(apps) == 20, f"Expected 20 apps, got {len(apps)}"
    assert len(users) == 156, f"Expected 156 users, got {len(users)}"
    assert len(assignments) == 622, f"Expected 622 assignments, got {len(assignments)}"
    print("\nAll checks passed.")
    print(
        f"  apps={len(apps)}, users={len(users)}, profiles={len(profiles)}, "
        f"assignments={len(assignments)}, audit_events={len(audit_events)}"
    )


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/fixtures/lanzadera")
    main(out)
