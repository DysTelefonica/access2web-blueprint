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
        raw = self.raw.strip()

        if "-" in raw:
            head, sep, tail = raw.partition("-")
        else:
            head, sep, tail = raw, "", ""

        if sep and tail == "":
            check_digit: str | None = None
        elif sep:
            check_digit = tail
        else:
            check_digit = None

        if not head.isdigit() or len(head) != 8:
            raise CPVCodeValidationError(f"CPV code must have exactly 8 digits, got {head!r}")

        if check_digit is not None and not check_digit.isdigit():
            raise CPVCodeValidationError(f"CPV check digit must be numeric, got {check_digit!r}")

        if check_digit is not None and len(check_digit) != 1:
            raise CPVCodeValidationError(
                f"CPV check digit must be exactly 1 digit, got {check_digit!r}"
            )

        object.__setattr__(self, "digits", head)
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
