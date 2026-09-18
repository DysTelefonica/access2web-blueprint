"""Commands and outcomes for EXP-CAP-034 canonical JSON."""

from dataclasses import dataclass
from typing import Any

CANONICAL_API_VERSION = "1.0"

CANONICAL_COLLECTIONS: tuple[str, ...] = (
    "expedientes",
    "hitos",
    "modificados",
    "anualidades",
    "anexos",
    "juridicas",
    "responsables",
    "suministradores",
    "lugares_ejecucion",
)


class JsonCanonicalError(Exception):
    """Base error for the canonical JSON use case."""


class JsonCanonicalValidationError(JsonCanonicalError):
    """The input collection set is invalid or the data is malformed."""


class JsonCanonicalAuthorizationError(JsonCanonicalError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class JsonCanonicalCommand:
    actor_id: Any
    collections: dict[str, list[dict[str, Any]]]
