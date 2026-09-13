"""Edit-expediente use case (CAP-002, issue #227).

CAP-002 §Camino feliz: "actualizan los campos manteniendo invariantes
y control de concurrencia". Optimistic locking via D-EXP-4: a stale
version returns ``ExpedienteEditConflictError`` (HTTP 409) without
commit.

DA-1: pure application layer — no framework imports.
"""

from app.src.modules.expedientes.application.edit_expediente.command import (
    ExpedienteEditAuthorizationError,
    ExpedienteEditCommand,
    ExpedienteEditConflictError,
    ExpedienteEditError,
    ExpedienteEditResult,
    ExpedienteEditValidationError,
)
from app.src.modules.expedientes.application.edit_expediente.service import (
    ExpedienteEditService,
)

__all__ = [
    "ExpedienteEditAuthorizationError",
    "ExpedienteEditCommand",
    "ExpedienteEditConflictError",
    "ExpedienteEditError",
    "ExpedienteEditResult",
    "ExpedienteEditService",
    "ExpedienteEditValidationError",
]
