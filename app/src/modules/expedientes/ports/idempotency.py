"""Transactional idempotency storage port (D-EXP-4)."""

from dataclasses import dataclass
from typing import Protocol
from uuid import UUID


@dataclass(frozen=True)
class IdempotencyRecord:
    key: UUID
    request_signature: str
    result: object


class IdempotencyPort(Protocol):
    async def get(self, key: UUID) -> IdempotencyRecord | None: ...
    async def record(self, record: IdempotencyRecord) -> None: ...
