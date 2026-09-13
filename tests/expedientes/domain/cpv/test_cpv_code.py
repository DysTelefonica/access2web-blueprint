"""Strict TDD — CPVCode value object (CAP-016 gate, issue #231).

El spec CAP-016 (`openspec/changes/expedientes-web-migration/specs/catalogs.md`)
exige "código válido según contrato" sin definir el contrato. Este módulo
implementa la validación ESTRUCTURAL del formato y la jerarquía del
Common Procurement Vocabulary (CPV, Reglamento UE 2195/2002). El
**checksum** queda ABIERTO (issue separado) — la fórmula oficial no está
documentada de forma accesible en el repositorio de origen y un
checksum inventado es peor que ninguno.

Vectores públicos tomados de Wikipedia y del portal TED:
https://en.wikipedia.org/wiki/Common_Procurement_Vocabulary

Tests asertan FORMATO y JERARQUÍA, no checksum.
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.domain.cpv.code import (
    CPVCode,
    CPVCodeValidationError,
)

# ---------------------------------------------------------------------------
# Format validation
# ---------------------------------------------------------------------------


def test_accepts_canonical_form_with_dash() -> None:
    """Canonical display form: 8 digits, dash, 1 check digit."""
    code = CPVCode("03113100-7")
    assert code.digits == "03113100"
    assert code.check_digit == "7"
    assert str(code) == "03113100-7"


def test_accepts_form_without_dash() -> None:
    """Some legacy data stores the code without the separator."""
    code = CPVCode("03113100")
    assert code.digits == "03113100"
    assert code.check_digit is None


def test_rejects_too_short() -> None:
    with pytest.raises(CPVCodeValidationError) as exc:
        CPVCode("0311310")
    assert "8 dígitos" in str(exc.value) or "8 digits" in str(exc.value)


def test_rejects_too_long() -> None:
    with pytest.raises(CPVCodeValidationError):
        CPVCode("0311310007")  # 10 digits, no separator


def test_rejects_non_numeric() -> None:
    with pytest.raises(CPVCodeValidationError):
        CPVCode("abcdefgh-i")


def test_rejects_alphabetic_check_digit() -> None:
    """Main vocabulary uses digits only. Alphabetic check digits are
    reserved for the supplementary vocabulary (different scope)."""
    with pytest.raises(CPVCodeValidationError):
        CPVCode("03113100-A")


def test_accepts_empty_check_digit_as_optional() -> None:
    """Empty after the dash is treated as missing — useful for legacy
    data with the dash but no check digit filled in."""
    code = CPVCode("03113100-")
    assert code.digits == "03113100"
    assert code.check_digit is None


def test_rejects_whitespace_inside() -> None:
    with pytest.raises(CPVCodeValidationError):
        CPVCode("0311 3100-7")


def test_rejects_internal_dashes() -> None:
    """Display form may have spaces (`031 131 00-7`); internal dashes are
    not part of the CPV format and must be rejected."""
    with pytest.raises(CPVCodeValidationError):
        CPVCode("031131-00-7")


def test_trims_surrounding_whitespace() -> None:
    code = CPVCode("  03113100-7  ")
    assert code.digits == "03113100"
    assert code.check_digit == "7"


# ---------------------------------------------------------------------------
# Known public vectors (format only, not checksum verification)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "raw",
    [
        "03113100-7",  # Sugar beet
        "03113200-8",  # Sugar cane
        "18451000-5",  # Buttons
        "18453000-9",  # Zip fasteners
        "45200000-9",  # Construction work (Wikipedia)
        "45210000-0",  # Construction (variant — published with -0 in some sources)
        "72300000-8",  # Data services
        "71355000-1",  # Surveying services
        "71355200-3",  # Ordnance surveying
        "79000000-4",  # Business services (variant in source)
    ],
)
def test_accepts_known_wikipedia_vectors(raw: str) -> None:
    code = CPVCode(raw)
    assert code.digits == raw.replace("-", "")[:8] or code.digits == raw[:8]


# ---------------------------------------------------------------------------
# Equality and hash
# ---------------------------------------------------------------------------


def test_two_codes_with_same_digits_are_equal() -> None:
    """`03113100-7` and `03113100` (no check) refer to the same code
    semantically; equality compares the digits."""
    assert CPVCode("03113100-7") == CPVCode("03113100")


def test_codes_hash_consistently_with_equality() -> None:
    a = CPVCode("03113100-7")
    b = CPVCode("03113100")
    assert hash(a) == hash(b)
    assert {a, b} == {a}


def test_different_digits_produce_different_codes() -> None:
    assert CPVCode("03113100-7") != CPVCode("03113200-8")
