"""Commands and outcomes for EXP-CAP-031 autosave related."""

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

from app.src.modules.expedientes.domain.anualidad import Anualidad
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.domain.modificado import Modificado


class AutosaveRelatedError(Exception):
    """Base error for related-data autosave."""


class AutosaveRelatedValidationError(AutosaveRelatedError):
    """The command or referenced expediente is invalid."""


class AutosaveRelatedAuthorizationError(AutosaveRelatedError):
    """The caller has no authenticated actor."""


class AutosaveRelatedConflictError(AutosaveRelatedError):
    """Version stale or idempotency key reused with different payload."""


@dataclass(frozen=True)
class AutosaveRelatedCommand:
    idempotency_key: UUID
    expediente_id: UUID
    expected_version: int
    hitos: tuple[Hito, ...]
    modificados: tuple[Modificado, ...]
    anualidades: tuple[Anualidad, ...]
    actor_id: UUID

    def signature(self) -> str:
        """Stable signature for idempotency replay detection."""
        ids = (
            tuple(sorted(str(h.id) for h in self.hitos))
            + tuple(sorted(str(m.id) for m in self.modificados))
            + tuple(sorted(str(a.id) for a in self.anualidades))
        )
        return f"{self.expediente_id}|{self.expected_version}|{'|'.join(ids)}"


@dataclass(frozen=True)
class AutosaveRelatedResult:
    expediente_id: UUID
    hitos_added: int
    modificados_added: int
    anualidades_added: int
    idempotency_key: UUID
    registered_at: datetime
