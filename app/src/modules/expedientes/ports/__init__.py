"""Expedientes ports layer — F01 (issue #222).

Purity rule (DA-1): ports may import ``domain`` and other ports only.
Frameworks are forbidden here.
"""

from __future__ import annotations

from app.src.modules.expedientes.ports.audit_log import AuditLogPort, ExpedienteAuditEvent
from app.src.modules.expedientes.ports.catalog_repository import CatalogRepositoryPort
from app.src.modules.expedientes.ports.document_storage import DocumentStoragePort, StoredDocument
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort
from app.src.modules.expedientes.ports.notification_delivery import NotificationDeliveryPort
from app.src.modules.expedientes.ports.readiness import (
    ReadinessCheck,
    ReadinessPort,
    ReadinessResult,
)

__all__ = [
    "AuditLogPort",
    "CatalogRepositoryPort",
    "DocumentStoragePort",
    "ExpedienteAuditEvent",
    "ExpedienteRepositoryPort",
    "HitoRepositoryPort",
    "NotificationDeliveryPort",
    "ReadinessCheck",
    "ReadinessPort",
    "ReadinessResult",
    "StoredDocument",
]
