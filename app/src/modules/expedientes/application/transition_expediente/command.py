"""ExpedienteTransitionCommand / Result / Error — CAP-004 boundary types.

CAP-004 §Camino feliz: "expediente elegible y padre válido cuando el
tipo lo requiere — se aplica el nuevo tipo y sus efectos derivados
de forma consistente, con salida determinista y auditable".

Este PR cubre CAP-004 (cambio de tipo + invariantes LOTE/BASED/padre).
CAP-005 (estado y garantía), CAP-006 (jerarquía acuerdos-lotes-basados)
y CAP-007 (ordinal) se cierran en follow-ups que extiendan F02 con los
campos ``fecha_fin_garantia``, ``titulo``, ``importe``, ``ordinal``.

§Validación y autorización: deniega sin efecto parcial.
§Concurrencia o fallo: preserva invariantes.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


class ExpedienteTransitionError(Exception):
    """Base error for the transition-expediente use case."""


class ExpedienteTransitionValidationError(ExpedienteTransitionError):
    """Malformed input, unknown aggregate, transition not in the
    matrix, or invariant violated (LOTE/BASED without padre, etc.).

    Maps to HTTP 400 in the delivery layer.
    """


class ExpedienteTransitionAuthorizationError(ExpedienteTransitionError):
    """The actor lacks permission to transition the aggregate.

    Maps to HTTP 403 in the delivery layer.
    """


class ExpedienteTransitionConflictError(ExpedienteTransitionError):
    """The ``expected_version`` doesn't match the stored version.

    D-EXP-4 (CAP-004 §Concurrencia): optimistic locking. Maps to HTTP 409
    in the delivery layer. The repository update is conditional on
    the version; zero affected rows means a stale version.
    """


@dataclass(frozen=True)
class ExpedienteTransitionCommand:
    """Inputs for CAP-004 cambio de tipo.

    ``new_tipo`` may equal ``None`` (no change) but the service then
    becomes a no-op and returns success without audit. ``None`` is
    allowed so the delivery layer can submit a single command for
    "I might want to change something" without knowing in advance
    whether the change is needed.

    ``new_id_expediente_padre`` may be ``None`` to leave the parent
    alone. Setting it to a UUID is required only when transitioning
    to LOTE or BASED (the aggregate's ``__post_init__`` enforces).
    """

    expediente_id: Any  # UUID
    expected_version: int
    new_tipo: Any = None  # ExpedienteTipo or None (no-op)
    new_id_expediente_padre: Any = None  # UUID or None (no change)
    motivo: str | None = None
    actor_id: Any = None  # UUID (required; service raises if None)


@dataclass(frozen=True)
class ExpedienteTransitionResult:
    """Outcome of a successful transition.

    ``new_tipo`` is the new type. ``new_id_expediente_padre`` is the
    new parent. ``new_version`` is the post-bump version
    (``expected_version + 1``). ``modified_at`` is the wall-clock
    instant the transition committed.
    """

    expediente_id: Any
    new_tipo: Any
    new_id_expediente_padre: Any
    new_version: int
    modified_at: datetime
