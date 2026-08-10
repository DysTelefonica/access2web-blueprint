# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — positive companion
"""A perfectly classifiable file, so the gate's failure can only come from its neighbour."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Expediente:
    identifier: int
