#!/usr/bin/env python3
# ci-pattern asset — detector único del 403 de plan del host (team-skills#330)
"""Un repositorio privado sin plan responde ``403`` («Upgrade to GitHub Pro…»)
al pedir la protección de rama o los rulesets del host. Ese cuerpo es EVIDENCIA
de que el host no puede aplicar esa área —no un error—, pero solo vale como
instantánea si el contrato del host declara el área ``unavailable`` en
``host_capabilities``.

Este módulo es el ÚNICO detector de ese 403: lo usan el gate de adopción
(``assets/bin/ci-pattern``, fase 3) y el readback del host
(``assets/host-readback/check_host_drift.py``), para que no existan dos
detectores que puedan divergir. Se importa con el directorio ``assets/`` en
``sys.path``::

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from host_plan import is_plan_403, capability_of
"""

from __future__ import annotations

PLAN_403_MARKER = "Upgrade to GitHub Pro"

# Áreas que declara `host_capabilities` en el contrato del host.
CAPABILITY_AREAS = ("branch_protection", "rulesets", "required_checks")

# Vocabulario cerrado de la capacidad. Un valor fuera de él no se puede
# interpretar: es DUDA (fail closed), nunca drift.
CAPABILITY_VALUES = frozenset({"available", "unavailable"})

# Área de `host_capabilities` que gobierna cada instantánea del contrato de
# adopción (`host_snapshots` de la fase 3).
SNAPSHOT_CAPABILITY = {
    "branch-protection": "branch_protection",
    "rulesets": "rulesets",
}


def is_plan_403(body):
    """True si ``body`` es el 403 de límite de plan del host.

    Se reconoce por el marcador del mensaje, no por el código de estado: el
    ``status`` puede venir como número o como cadena según quién escriba la
    instantánea, y el marcador es el que identifica la causa.
    """
    return (isinstance(body, dict)
            and isinstance(body.get("message"), str)
            and PLAN_403_MARKER in body["message"])


def capability_of(contract, area):
    """Capacidad declarada para un área, o ``None`` si no se declara.

    Un contrato ausente, sin ``host_capabilities`` o con un valor que no es
    cadena deja el área SIN declarar: quien decide entonces no puede
    interpretar el 403 y debe fallar cerrado.
    """
    if not isinstance(contract, dict):
        return None
    caps = contract.get("host_capabilities")
    if not isinstance(caps, dict):
        return None
    value = caps.get(area)
    return value if isinstance(value, str) and value.strip() else None
