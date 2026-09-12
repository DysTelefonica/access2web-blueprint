"""CPV hierarchy parsing (CAP-016, Reglamento UE 2195/2002 Anexo I).

El CPV principal es una jerarquía de cinco niveles:

- División  (XX000000)  — los dos primeros dígitos
- Grupo     (XXX00000)  — los tres primeros
- Clase     (XXXX0000)  — los cuatro primeros
- Categoría (XXXXX000)  — los cinco primeros
- Subcategoría (XXXXXXXX) — los ocho

Un prefijo de N dígitos representa el nivel N. Estos helpers operan
sobre strings de dígitos directamente (no sobre ``CPVCode``) porque
las queries jerárquicas —"todo lo que empieza por 03"— necesitan
operar sobre prefijos arbitrarios sin tener que construir un
``CPVCode`` válido primero.

Este módulo NO decide si un código existe en el catálogo (eso lo hace
el adapter con el maestro cargado); solo extrae la jerarquía del
prefijo.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from enum import StrEnum


class HierarchyLevel(StrEnum):
    """Position in the CPV main vocabulary tree.

    StrEnum so the level survives JSON serialisation (audit logs,
    API responses) and round-trips through environment variables.
    """

    UNCLASSIFIED = "unclassified"
    DIVISION = "division"
    GROUP = "group"
    CLASS = "class"
    CATEGORY = "category"
    SUBCATEGORY = "subcategory"


def hierarchy_level(digits: str) -> HierarchyLevel:
    """Identify the hierarchy level represented by the digit prefix.

    The level is determined by the number of digits:

    - 0–1 digits → UNCLASSIFIED
    - 2 digits   → DIVISION
    - 3 digits   → GROUP
    - 4 digits   → CLASS
    - 5 digits   → CATEGORY
    - 6+ digits  → SUBCATEGORY (the canonical 8-digit form)

    Inputs longer than 8 are accepted but only the first 8 digits
    are considered. ``CPVCode.digits`` already enforces exactly 8
    for canonical codes; the leniency here is for query helpers.
    """
    if len(digits) >= 6:
        return HierarchyLevel.SUBCATEGORY
    if len(digits) == 5:
        return HierarchyLevel.CATEGORY
    if len(digits) == 4:
        return HierarchyLevel.CLASS
    if len(digits) == 3:
        return HierarchyLevel.GROUP
    if len(digits) == 2:
        return HierarchyLevel.DIVISION
    return HierarchyLevel.UNCLASSIFIED


def division(digits: str) -> str:
    """Return the division prefix (first two digits)."""
    return digits[:2]


def group(digits: str) -> str:
    """Return the group prefix (first three digits)."""
    return digits[:3]


def klass(digits: str) -> str:
    """Return the class prefix (first four digits).

    Named ``klass`` (not ``class``) because ``class`` is a Python
    keyword. The CPV term is the English noun "class".
    """
    return digits[:4]


def category(digits: str) -> str:
    """Return the category prefix (first five digits)."""
    return digits[:5]
