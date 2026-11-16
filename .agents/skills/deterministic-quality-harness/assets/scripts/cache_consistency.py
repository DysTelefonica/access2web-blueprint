#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v2.3 — assets/scripts/cache_consistency.py
"""Read-only consistency check for a derived (denormalized) cache (Hard Rule 20).

The consuming repo supplies three reads from one read-only snapshot: the live row keys, the cached
value per key, and the source derivation per key. This module only compares; it never writes the
cache. Pattern, write paths that forget to recompute, and the single-issue job:
`references/derived-cache-consistency.md`.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass

EXIT_OK = 0
EXIT_DRIFT = 1


@dataclass(frozen=True)
class Drift:
    """One live row whose cached value differs from the source derivation."""

    key: str
    cached: str | None
    derived: str

    @property
    def kind(self) -> str:
        """``missing`` when the cache has no row for a live key, ``mismatch`` otherwise."""
        return "missing" if self.cached is None else "mismatch"


def compare(
    live_keys: Iterable[str],
    cached: Mapping[str, str],
    derive: Callable[[str], str],
) -> list[Drift]:
    """Return every live row whose cached value is absent or differs from ``derive(key)``."""
    drift: list[Drift] = []
    for key in live_keys:
        derived = derive(key)
        value = cached.get(key)
        if value != derived:
            drift.append(Drift(key=key, cached=value, derived=derived))
    return drift


def exit_code(drift: list[Drift]) -> int:
    """``0`` when the cache matches its source, ``1`` when any live row drifted."""
    return EXIT_DRIFT if drift else EXIT_OK


def summary(drift: list[Drift], sample_size: int = 20) -> str:
    """Issue-ready text: the full count plus a bounded sample of differing rows."""
    lines = [f"differences: {len(drift)}"]
    lines.extend(f"{d.key}\t{d.kind}\tcached={d.cached!r}\tderived={d.derived!r}" for d in drift[:sample_size])
    return "\n".join(lines)
