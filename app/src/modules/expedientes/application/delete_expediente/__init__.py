"""Delete-expediente use case (CAP-003, issue #228).

CAP-003 §Camino feliz: "se elimina el agregado de forma auditable
sin afectar relaciones no incluidas". Conditional aspect: by default
the use case refuses when the aggregate has children; ``force=True``
in the command overrides.

DA-1: pure application layer — no framework imports.
"""

from app.src.modules.expedientes.application.delete_expediente.command import (
    ExpedienteDeleteAuthorizationError,
    ExpedienteDeleteCommand,
    ExpedienteDeleteConflictError,
    ExpedienteDeleteError,
    ExpedienteDeleteResult,
    ExpedienteDeleteValidationError,
)
from app.src.modules.expedientes.application.delete_expediente.service import (
    ExpedienteDeleteService,
)

__all__ = [
    "ExpedienteDeleteAuthorizationError",
    "ExpedienteDeleteCommand",
    "ExpedienteDeleteConflictError",
    "ExpedienteDeleteError",
    "ExpedienteDeleteResult",
    "ExpedienteDeleteService",
    "ExpedienteDeleteValidationError",
]
