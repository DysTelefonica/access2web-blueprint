"""Strict TDD — CPV hierarchy parsing (CAP-016 gate).

El CPV es jerárquico: División (XX), Grupo (XXX), Clase (XXXX),
Categoría (XXXXX). El código completo (XXXXXXXX) añade dos dígitos
de granularidad fina. Estos tests verifican que la jerarquía se
extrae correctamente a partir del código completo.

FUENTE: Reglamento (CE) 2195/2002 Anexo I, §Estructura del sistema
de clasificación.

Wikipedia lo confirma (párrafos "Divisions: first two digits...").

Nota: las funciones de jerarquía operan sobre strings de 8 dígitos
(el `digits` extraído de un ``CPVCode``), NO sobre ``CPVCode``
directamente. Los prefijos jerárquicos no son códigos CPV válidos
por sí mismos; son herramientas para queries tipo "todo lo que
empieza por 03".
"""

from __future__ import annotations

from app.src.modules.expedientes.domain.cpv.code import CPVCode
from app.src.modules.expedientes.domain.cpv.structure import (
    HierarchyLevel,
    category,
    division,
    group,
    hierarchy_level,
    klass,
)

# ---------------------------------------------------------------------------
# Hierarchy level detection (operates on a digits string)
# ---------------------------------------------------------------------------


def test_eight_digit_digits_is_subcategory() -> None:
    """A full 8-digit CPV digits (or 9 with check) is the most granular level."""
    assert hierarchy_level(CPVCode("03113100-7").digits) is HierarchyLevel.SUBCATEGORY


def test_eight_digit_digits_without_check_is_subcategory() -> None:
    assert hierarchy_level(CPVCode("03113100").digits) is HierarchyLevel.SUBCATEGORY


def test_empty_digits_is_unclassified() -> None:
    assert hierarchy_level("") is HierarchyLevel.UNCLASSIFIED


def test_one_digit_is_unclassified() -> None:
    assert hierarchy_level("0") is HierarchyLevel.UNCLASSIFIED


def test_nine_plus_digit_is_subcategory() -> None:
    """The function operates on the first 8 digits; longer inputs are
    truncated (and the caller is responsible for using ``CPVCode.digits``
    which already enforces 8 digits)."""
    assert hierarchy_level("031131001") is HierarchyLevel.SUBCATEGORY


# ---------------------------------------------------------------------------
# Hierarchical extraction
# ---------------------------------------------------------------------------


def test_division_is_first_two_digits() -> None:
    assert division(CPVCode("03113100-7").digits) == "03"


def test_group_is_first_three_digits() -> None:
    assert group(CPVCode("03113100-7").digits) == "031"


def test_class_is_first_four_digits() -> None:
    assert klass(CPVCode("03113100-7").digits) == "0311"


def test_category_is_first_five_digits() -> None:
    assert category(CPVCode("03113100-7").digits) == "03113"


# ---------------------------------------------------------------------------
# Hierarchical containment
# ---------------------------------------------------------------------------


def test_codes_in_same_division_share_division() -> None:
    assert division(CPVCode("03113100-7").digits) == division(CPVCode("03999999-9").digits)


def test_codes_in_different_divisions_do_not_share_division() -> None:
    assert division(CPVCode("03113100-7").digits) != division(CPVCode("18131000-5").digits)


# ---------------------------------------------------------------------------
# Pure-string usage (queries need to slice without constructing CPVCode)
# ---------------------------------------------------------------------------


def test_helpers_work_on_raw_eight_digit_strings() -> None:
    """Query builders often have a digit string (e.g. from a search
    input) and need to compute the division without round-tripping
    through ``CPVCode``. The helpers accept strings directly.
    """
    assert division("03113100") == "03"
    assert group("03113100") == "031"
    assert klass("03113100") == "0311"
    assert category("03113100") == "03113"
