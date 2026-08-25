# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Shared port imports for the admin HTTP delivery modules.

W47 (#494) split ``admin.py`` into four modules. Three of them
(``admin.py``, ``admin_routes.py``, plus the route files) carry
the same five port imports. Centralising them here avoids the
``dup:5653b58ccd4b`` DRY dup that would otherwise appear across
``admin.py`` and ``admin_routes.py``.
"""

from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)

__all__ = [
    "AppRepositoryPort",
    "AssignmentRepositoryPort",
    "AuditLogPort",
    "GlobalAdminRepositoryPort",
    "UserRepository",
]
