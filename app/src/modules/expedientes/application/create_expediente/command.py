"""ExpedienteAltaCommand / Result / Error — CAP-001 boundary types.

CAP-001 §Camino feliz pide "se persisten cabecera, hijos, read-model y
último cambio en una transacción confirmada, con salida
determinista y auditable". El command captura los inputs del
alta; el Result captura el resultado. La taxonomía de errores
separa los fallos de validación (400) de los fallos de runtime
(500/409) para que el delivery layer pueda mapearlos sin
re-inspeccionar el mensaje.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


class ExpedienteAltaError(Exception):
    """Base error for the create-expediente use case.

    Subclass for the specific failure mode so the HTTP delivery
    layer (CAP-001 §Validación y autorización) can return the right
    status code without re-inspecting the message.
    """


class ExpedienteAltaValidationError(ExpedienteAltaError):
    """The input does not satisfy the domain invariants.

    Examples: empty codigo_ordinal, LOTE without id_expediente_padre.
    Maps to HTTP 400 in the delivery layer.
    """


class ExpedienteAltaAuthorizationError(ExpedienteAltaError):
    """The actor lacks permission to create the expediente.

    Maps to HTTP 403 in the delivery layer.
    """


class ExpedienteAltaConflictError(ExpedienteAltaError):
    """The codigo_ordinal is already taken (CAP-001 §Concurrencia).

    Maps to HTTP 409 in the delivery layer.
    """


@dataclass(frozen=True)
class ExpedienteAltaCommand:
    """Inputs for CAP-001 alta transaccional.

    The fields mirror the legacy ``TbExpedientes`` (data-model.md §14)
    plus the actor for audit (D-EXP-3 deny-by-default requires the
    use case to record who initiated the alta).
    """

    codigo_ordinal: str
    tipo: Any  # ExpedienteTipo enum; kept as Any to avoid domain import cycle
    actor_id: Any  # UUID
    id_expediente_padre: Any = None  # UUID or None for AM


@dataclass(frozen=True)
class ExpedienteAltaResult:
    """Outcome of a successful alta.

    ``expediente_id`` is the new aggregate identity; ``version`` is
    the optimistic-locking counter (always 1 for a new aggregate —
    CAP-001 implies version starts at 1 per D-EXP-4); ``created_at``
    is the wall-clock instant the alta committed.
    """

    expediente_id: Any  # UUID
    version: int
    created_at: datetime
