"""Create-expediente use case (CAP-001, issue #226).

Vertical lifecycle de alta transaccional. Sigue el patrón
``application/<use-case>/{command,service}`` establecido por la
arquitectura hexagonal del módulo.

DA-1: pure application layer — no framework imports.
"""

from app.src.modules.expedientes.application.create_expediente.command import (
    ExpedienteAltaAuthorizationError,
    ExpedienteAltaCommand,
    ExpedienteAltaConflictError,
    ExpedienteAltaError,
    ExpedienteAltaResult,
    ExpedienteAltaValidationError,
)
from app.src.modules.expedientes.application.create_expediente.service import (
    ExpedienteAltaService,
)

__all__ = [
    "ExpedienteAltaAuthorizationError",
    "ExpedienteAltaCommand",
    "ExpedienteAltaConflictError",
    "ExpedienteAltaError",
    "ExpedienteAltaResult",
    "ExpedienteAltaService",
    "ExpedienteAltaValidationError",
]
