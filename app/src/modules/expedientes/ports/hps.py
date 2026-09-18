"""HPS integration port (CAP-051, H01).

Duck-typed Protocol to keep cross-module application→driver imports out
of the application layer (see ``scripts/check_layers.py``).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol
from uuid import UUID


@dataclass(frozen=True)
class HpsSubmission:
    """One HPS submission contract (CAP-051 camino feliz)."""

    idempotency_key: UUID
    expediente_id: UUID
    payload: dict[str, Any]


@dataclass(frozen=True)
class HpsRecord:
    """One HPS record retrieved by reference (CAP-051 consulta)."""

    reference: UUID
    status: str
    payload: dict[str, Any]


class HpsPort(Protocol):
    """Authenticates, deduplicates by ``idempotency_key``, and emits HPS records.

    The concrete adapter is injected by the DI container; tests use a
    dataclass fake with ``async def submit`` and ``async def query``.
    """

    async def submit(self, submission: HpsSubmission, credential: str) -> HpsRecord: ...
    async def query(self, reference: UUID, credential: str) -> HpsRecord: ...


__all__ = ["HpsPort", "HpsRecord", "HpsSubmission"]
