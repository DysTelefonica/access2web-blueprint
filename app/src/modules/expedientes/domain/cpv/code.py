"""CPVCode value object — format validation only.

El CPV (Common Procurement Vocabulary, Reglamento UE 2195/2002)
es un código jerárquico de 8 dígitos + 1 check digit. Este módulo
valida el FORMATO del código (longitud, separador, caracteres
numéricos, opcionalidad del check digit).

El **checksum** del check digit queda ABIERTO: la fórmula oficial
no está documentada de forma accesible en el repositorio de origen
del CPV y un checksum inventado sería peor que ninguno. El puerto
``CPVCodeValidator`` permite inyectar un validador con checksum real
cuando la fórmula esté documentada (issue de seguimiento).

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass


class CPVCodeValidationError(ValueError):
    """Raised when a string is not a syntactically valid CPV code.

    The error covers format only. Semantic validation (does this code
    exist in the master catalogue?) is the adapter's job.
    """


def _split_digits_and_check(raw: str) -> tuple[str, str | None]:
    """Split ``raw`` into the 8-digit body and an optional check digit.

    Returns ``(body, check_digit)``. Raises :class:`CPVCodeValidationError`
    on malformed input.
    """
    body, sep, tail = raw.partition("-")
    if not sep:
        return body, None
    if tail == "":
        return body, None
    return body, tail


def _validate_body(body: str) -> None:
    if not body.isdigit() or len(body) != 8:
        raise CPVCodeValidationError(
            f"CPV code must have exactly 8 digits, got {body!r}"
        )


def _validate_check_digit(check_digit: str | None) -> None:
    if check_digit is None:
        return
    if not check_digit.isdigit() or len(check_digit) != 1:
        raise CPVCodeValidationError(
            f"CPV check digit must be exactly 1 numeric digit, got {check_digit!r}"
        )


@dataclass(frozen=True)
class CPVCode:
    """A syntactically valid CPV code.

    Construction takes the raw textual form. Two forms are accepted:

    - ``03113100-7`` — canonical display form (8 digits, dash, check).
    - ``03113100`` — storage form without check digit.

    Surrounding whitespace is trimmed. Internal whitespace and extra
    dashes are rejected so typos surface at parse time, not silently
    in the database.

    Equality compares the eight digits only: ``03113100-7`` and
    ``03113100`` are the same code. The check digit is kept as
    metadata for audit and traceability.
    """

    raw: str
    digits: str = ""
    check_digit: str | None = None

    def __post_init__(self) -> None:
        body, check_digit = _split_digits_and_check(self.raw.strip())
        _validate_body(body)
        _validate_check_digit(check_digit)

        object.__setattr__(self, "digits", body)
        object.__setattr__(self, "check_digit", check_digit)
        object.__setattr__(self, "raw", self.__str__())

    def __str__(self) -> str:
        if self.check_digit is None:
            return self.digits
        return f"{self.digits}-{self.check_digit}"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CPVCode):
            return NotImplemented
        return self.digits == other.digits

    def __hash__(self) -> int:
        return hash(self.digits)
