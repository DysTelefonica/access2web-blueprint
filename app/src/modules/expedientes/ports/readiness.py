"""ReadinessPort — CAP-047, D-EXP-3."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Protocol


@dataclass
class ReadinessCheck:
    name: str
    ok: bool
    message: str = ""


@dataclass
class ReadinessResult:
    ready: bool
    checks: list[ReadinessCheck] = field(default_factory=list)


class ReadinessPort(Protocol):
    """El runtime interroga este puerto antes de declarar ``ready``.

    D-EXP-3: deniegue sin stub permisivo en producción.
    """

    async def check(self) -> ReadinessResult: ...
