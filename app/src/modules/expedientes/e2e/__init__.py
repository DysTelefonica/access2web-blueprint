"""Expedientes E2E — vertical de sincronización y hash (CAP-036+).

El paquete agrupa la lógica de canonización, hash versionado y
generación de paquetes E2E con el sistema externo. Sigue D-EXP-7
(Canonicalizador versionado) y se relaciona con CAP-036 (Hash versionado),
CAP-037 (Paquete E2E), CAP-042 (Ordinal E2E).

DA-1: pure domain — no framework imports.
"""

from app.src.modules.expedientes.e2e.hash_versioning import (
    FNV1A32,
    HashAlgorithm,
    HashRegistry,
    HashVersioned,
    canonicalize,
    compute_hash,
    default_registry,
)

__all__ = [
    "FNV1A32",
    "HashAlgorithm",
    "HashRegistry",
    "HashVersioned",
    "canonicalize",
    "compute_hash",
    "default_registry",
]
