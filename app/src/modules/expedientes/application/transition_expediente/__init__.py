"""Transition-expediente use case (CAP-004, issue #229).

CAP-004 §Camino feliz: "se aplica el nuevo tipo y sus efectos
derivados de forma consistente, con salida determinista y auditable".
The transition matrix is the source of truth for valid transitions;
the service enforces it before opening the UoW (fail-fast).

DA-1: pure application layer — no framework imports.
"""

from app.src.modules.expedientes.application.transition_expediente.command import (
    ExpedienteTransitionAuthorizationError,
    ExpedienteTransitionCommand,
    ExpedienteTransitionConflictError,
    ExpedienteTransitionError,
    ExpedienteTransitionResult,
    ExpedienteTransitionValidationError,
)
from app.src.modules.expedientes.application.transition_expediente.service import (
    ALLOWED_TRANSITIONS,
    ExpedienteTransitionService,
)

__all__ = [
    "ALLOWED_TRANSITIONS",
    "ExpedienteTransitionAuthorizationError",
    "ExpedienteTransitionCommand",
    "ExpedienteTransitionConflictError",
    "ExpedienteTransitionError",
    "ExpedienteTransitionResult",
    "ExpedienteTransitionService",
    "ExpedienteTransitionValidationError",
]
