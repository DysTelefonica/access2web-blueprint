"""EXP-CAP-046 Auditoría adapter (A04)."""

from app.src.modules.expedientes.application.audit.command import (
    AuditAuthorizationError,
    AuditCommand,
    AuditDependencyError,
    AuditError,
    AuditResult,
)
from app.src.modules.expedientes.application.audit.service import AuditService

__all__ = [
    "AuditAuthorizationError",
    "AuditCommand",
    "AuditDependencyError",
    "AuditError",
    "AuditResult",
    "AuditService",
]
