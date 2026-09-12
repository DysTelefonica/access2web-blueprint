"""Anexos domain — referencias documentales opacas (CAP-010, CAP-056).

Este paquete encapsula los value objects de anexos:

- ``AnexoReference``: referencia opaca + MIME type + size (CAP-010 §Camino
  feliz: "se registra la referencia ... sin exponer rutas locales").
- ``SizeLimitPolicy``: límite de tamaño configurable con verificación
  ruidosa (D-EXP-3 deny-by-default; el valor concreto queda como
  decisión de producto, see #253 H06).
- ``RetentionPolicy``: política de retención por niveles (D28) con
  ``RetentionLevel.OPERATIONAL|LEGAL|HISTORICAL``. Cálculo determinista
  de ``expires_at`` y ``is_expired``.

Los anexos viven tras ``DocumentStoragePort``
(``app/src/modules/expedientes/ports/document_storage.py``), que ya
provee la interfaz para el adapter. Lo que este paquete añade es el
modelo de dominio que el spec exige (referencia opaca + límite +
retención).

DA-1: pure domain — no framework imports.
"""

from app.src.modules.expedientes.domain.anexo.reference import (
    AnexoReference,
    AnexoReferenceValidationError,
)
from app.src.modules.expedientes.domain.anexo.retention import (
    RetentionExpired,
    RetentionLevel,
    RetentionPolicy,
)
from app.src.modules.expedientes.domain.anexo.size_limit import (
    DEFAULT_MAX_SIZE_BYTES,
    SizeLimitExceeded,
    SizeLimitPolicy,
)

__all__ = [
    "AnexoReference",
    "AnexoReferenceValidationError",
    "DEFAULT_MAX_SIZE_BYTES",
    "RetentionExpired",
    "RetentionLevel",
    "RetentionPolicy",
    "SizeLimitExceeded",
    "SizeLimitPolicy",
]
