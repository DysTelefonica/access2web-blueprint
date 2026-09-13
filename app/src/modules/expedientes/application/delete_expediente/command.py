"""ExpedienteDeleteCommand / Result / Error — CAP-003 boundary types.

CAP-003 §Camino feliz exige "se elimina el agregado de forma
auditable sin afectar relaciones no incluidas". §Validación y
autorización: "deniega sin efecto parcial, devuelve error
diagnosticable y conserva auditoría". §Concurrencia o fallo:
"preserva invariantes, no duplica ni pierde evidencia, queda
reintentable".

El flag ``force`` implementa "impedir pérdida de hijos": el use case
rehúsa por defecto si el agregado tiene hijos, a menos que el
caller pase ``force=True`` y acepte el riesgo.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


class ExpedienteDeleteError(Exception):
    """Base error for the delete-expediente use case."""


class ExpedienteDeleteValidationError(ExpedienteDeleteError):
    """Malformed input, unknown aggregate, or conditional-delete
    refused because the aggregate has children.

    Maps to HTTP 400 in the delivery layer.
    """


class ExpedienteDeleteAuthorizationError(ExpedienteDeleteError):
    """The actor lacks permission to delete the aggregate.

    Maps to HTTP 403 in the delivery layer.
    """


class ExpedienteDeleteConflictError(ExpedienteDeleteError):
    """The ``expected_version`` doesn't match the stored version.

    D-EXP-4 (CAP-003 §Concurrencia): optimistic locking. Maps to HTTP 409
    in the delivery layer. The repository delete is conditional on
    the version (``DELETE ... WHERE version = ?``); zero affected rows
    means a stale version.
    """


@dataclass(frozen=True)
class ExpedienteDeleteCommand:
    """Inputs for CAP-003 baja condicionada.

    ``force`` bypasses the children-presence check. Default False:
    refusing children preserves CAP-003 "impedir pérdida de hijos".
    """

    expediente_id: Any  # UUID
    expected_version: int
    force: bool = False
    actor_id: Any = None  # UUID (required; service raises if None)


@dataclass(frozen=True)
class ExpedienteDeleteResult:
    """Outcome of a successful delete.

    ``deleted_at`` is the wall-clock instant the delete committed.
    """

    expediente_id: Any
    deleted_at: datetime
