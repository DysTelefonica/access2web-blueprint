"""ExpedienteEditCommand / Result / Error — CAP-002 boundary types.

CAP-002 §Camino feliz exige "actualizan los campos manteniendo
invariantes y control de concurrencia". El command captura los
campos opcionales a editar; ``None`` significa "no cambiar". La
taxonomía de errores separa los fallos de validación (400), los
conflictos de versión (409), la autorización (403) y los errores de
runtime (500) para que el delivery layer los mapee a HTTP status
codes sin re-inspeccionar el mensaje.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


class ExpedienteEditError(Exception):
    """Base error for the edit-expediente use case."""


class ExpedienteEditValidationError(ExpedienteEditError):
    """Malformed input or unknown aggregate.

    Maps to HTTP 400 in the delivery layer.
    """


class ExpedienteEditAuthorizationError(ExpedienteEditError):
    """The actor lacks permission to edit the aggregate.

    Maps to HTTP 403 in the delivery layer.
    """


class ExpedienteEditConflictError(ExpedienteEditError):
    """The ``expected_version`` doesn't match the stored version.

    D-EXP-4 (CAP-002 §Concurrencia): optimistic locking. Maps to HTTP 409
    in the delivery layer. The repository update is conditional
    (UPDATE ... WHERE version = ?) so this exception means zero rows
    were affected.
    """


@dataclass(frozen=True)
class ExpedienteEditCommand:
    """Inputs for CAP-002 edición concurrente.

    Fields marked ``Optional`` mean "leave unchanged if None". This
    shape lets the caller send only the fields it wants to change.

    The current field surface is intentionally small: only
    ``estado`` (state transitions) and ``id_expediente_padre`` (move
    to a different parent). Other fields (titulo, importe, fechas,
    ordinal) require the F02 aggregate to carry them — that gap is
    documented in the PR body for C02.
    """

    expediente_id: Any  # UUID
    expected_version: int
    estado: Any | None = None  # ExpedienteEstado enum or None
    id_expediente_padre: Any | None = None  # UUID or None
    actor_id: Any = None  # UUID (required; service raises if None)


@dataclass(frozen=True)
class ExpedienteEditResult:
    """Outcome of a successful edit.

    ``new_version`` is the aggregate's version after the bump
    (always ``expected_version + 1`` on success). ``modified_at``
    is the wall-clock instant the edit committed.
    """

    expediente_id: Any
    new_version: int
    modified_at: datetime
    fields_changed: tuple[str, ...]
