"""Anexo size limit policy (CAP-010, D-EXP-3).

CAP-010 §Camino feliz: "dentro del límite acordado". El límite
concreto es decisión de producto (issue #253 H06); el módulo entrega
una política configurable con un valor por defecto conservador
(25 MiB) y un verificador que produce un error diagnosticable.

D-EXP-3 (deny-by-default) implica que un anexo fuera del límite
debe rechazarse ruidosamente, no truncarse ni aceptarse en
silencio. ``SizeLimitExceeded`` lleva el contexto de la operación
para los logs de auditoría.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass

# 25 MiB = 25 * 1024 * 1024 bytes. Spanish public-sector IT shops
# commonly use 25 MB as the upper bound for anexos; until product
# decides otherwise, 25 MiB keeps storage costs predictable and
# matches the spec's open question (design.md D-63).
DEFAULT_MAX_SIZE_BYTES: int = 25 * 1024 * 1024


class SizeLimitExceeded(ValueError):
    """Raised when an anexo exceeds the configured size limit.

    Carries the offending size, the configured limit, and the
    operation context for the audit log.
    """

    def __init__(self, size_bytes: int, limit_bytes: int, context: str) -> None:
        self.size_bytes = size_bytes
        self.limit_bytes = limit_bytes
        self.context = context
        super().__init__(
            f"anexo size {size_bytes} bytes exceeds limit "
            f"{limit_bytes} bytes ({_format_mib(limit_bytes)}) in {context!r}"
        )


def _format_mib(num_bytes: int) -> str:
    """Format a byte count as a human-readable MiB string."""
    mib = num_bytes / (1024 * 1024)
    return f"{mib:.1f} MiB" if mib >= 1 else f"{num_bytes} B"


@dataclass(frozen=True)
class SizeLimitPolicy:
    """A configurable upper bound on anexo sizes.

    Construct with the agreed limit (in bytes). ``ensure_within_limit``
    is the runtime check that produces a ``SizeLimitExceeded`` with the
    right context for the audit log.
    """

    max_size_bytes: int

    def __post_init__(self) -> None:
        if self.max_size_bytes < 0:
            raise ValueError(f"max_size_bytes must be non-negative, got {self.max_size_bytes}")

    def ensure_within_limit(self, *, size_bytes: int, context: str) -> None:
        """Raise ``SizeLimitExceeded`` if ``size_bytes`` exceeds the limit.

        Zero and negative sizes are rejected with ``ValueError``: a
        zero-size anexo is malformed input (not an annexable artefact),
        and a negative size is a programming error elsewhere (a size
        should never be negative). Both are caught at this gate so the
        downstream adapter does not have to defend against them.
        """
        if size_bytes <= 0:
            raise ValueError(f"size_bytes must be positive, got {size_bytes}")
        if size_bytes > self.max_size_bytes:
            raise SizeLimitExceeded(
                size_bytes=size_bytes,
                limit_bytes=self.max_size_bytes,
                context=context,
            )
