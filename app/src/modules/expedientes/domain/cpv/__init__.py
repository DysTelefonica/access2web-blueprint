"""CPV domain — value object y jerarquía (CAP-016 gate).

Este paquete encapsula el formato y la jerarquía del Common Procurement
Vocabulary. La validación semántica (¿existe este código en la maestra
del dominio?) vive en el adapter y se inyecta vía ``CatalogRepositoryPort``
(F01). El checksum oficial queda ABIERTO (issue de seguimiento).

DA-1: pure domain — no framework imports.
"""

from app.src.modules.expedientes.domain.cpv.code import CPVCode, CPVCodeValidationError
from app.src.modules.expedientes.domain.cpv.structure import (
    HierarchyLevel,
    category,
    division,
    group,
    hierarchy_level,
    klass,
)

__all__ = [
    "CPVCode",
    "CPVCodeValidationError",
    "HierarchyLevel",
    "category",
    "division",
    "group",
    "hierarchy_level",
    "klass",
]
