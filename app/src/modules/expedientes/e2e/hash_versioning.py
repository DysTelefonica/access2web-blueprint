"""Hash versionado y canonicalización (CAP-036).

El spec `openspec/changes/expedientes-web-migration/specs/e2e.md`
exige una estrategia de hash versionada y reproducible. Este módulo
implementa la infraestructura neutral que sirve para CUALQUIER
política de hash (FNV-1a u otra). La elección del algoritmo concreto
es decisión humana (issue #659 — FNV-1a golden o nuevo algoritmo).

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Protocol

# ---------------------------------------------------------------------------
# Canonicalización
# ---------------------------------------------------------------------------


def canonicalize(payload: object) -> bytes:
    """Serialize ``payload`` to a canonical UTF-8 byte sequence.

    The canonical form is order-independent on the producer side: a
    dict and a dict with the same entries in a different order produce
    the same bytes. Whitespace is fixed (``separators=(",", ":")``)
    so the output is byte-stable across Python versions and hosts.

    The function is recursive: nested dicts and lists are normalised
    at every level. Values that are not JSON-serialisable (sets,
    custom classes) raise ``TypeError`` so the producer can fix the
    payload instead of relying on Python's defaults.
    """
    return json.dumps(
        payload,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
        allow_nan=False,
    ).encode("utf-8")


# ---------------------------------------------------------------------------
# Algorithm interface and registry
# ---------------------------------------------------------------------------


class HashAlgorithm(Protocol):
    """A hash algorithm the registry can dispatch to."""

    @property
    def algorithm(self) -> str: ...

    def hash(self, payload: bytes) -> str:
        """Return the lowercase hex digest of ``payload``."""
        ...


class HashRegistry:
    """In-memory registry of hash algorithms.

    Resolution is by string name (e.g. ``"fnv1a-32"``). Registrations are
    immutable once added; double-registration of the same name is a
    hard error so the producer cannot silently swap algorithms
    underneath an in-flight payload.
    """

    def __init__(self) -> None:
        self._algorithms: dict[str, HashAlgorithm] = {}

    def register(self, algorithm: HashAlgorithm) -> None:
        if algorithm.algorithm in self._algorithms:
            raise ValueError(f"hash algorithm {algorithm.algorithm!r} already registered")
        self._algorithms[algorithm.algorithm] = algorithm

    def get(self, name: str) -> HashAlgorithm:
        try:
            return self._algorithms[name]
        except KeyError as exc:
            raise KeyError(
                f"hash algorithm {name!r} not in registry (known: {sorted(self._algorithms)})"
            ) from exc

    def names(self) -> list[str]:
        return sorted(self._algorithms)


default_registry = HashRegistry()


# ---------------------------------------------------------------------------
# Hash result and top-level API
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class HashVersioned:
    """A hash result that carries the algorithm that produced it.

    CAP-036: the strategy MUST be versioned. The ``algorithm`` field
    is the version stamp: a reader that sees ``"fnv1a-32"`` knows
    how to reproduce the digest; a future reader that sees
    ``"sha256-v2"`` knows to switch to a different algorithm.
    """

    algorithm: str
    digest: str


def compute_hash(
    payload: object,
    *,
    algorithm: str,
    registry: HashRegistry | None = None,
) -> HashVersioned:
    """Compute a versioned hash for ``payload`` using ``algorithm``.

    Returns a ``HashVersioned`` with the algorithm name and the
    lowercase hex digest. Raises ``KeyError`` if the algorithm is
    not registered (see :class:`HashRegistry`).
    """
    registry = registry if registry is not None else default_registry
    algo = registry.get(algorithm)
    canonical = canonicalize(payload)
    digest = algo.hash(canonical)
    return HashVersioned(algorithm=algo.algorithm, digest=digest)


# ---------------------------------------------------------------------------
# FNV-1a 32-bit (default algorithm — provisional, see issue #659)
# ---------------------------------------------------------------------------


_FNV1A_32_OFFSET_BASIS = 0x811C9DC5
_FNV1A_32_PRIME = 0x01000193


class FNV1A32:
    """FNV-1a 32-bit hash.

    Provisional default in :data:`default_registry`. Decision on
    whether to keep FNV-1a or switch to another algorithm is tracked
    in issue #659. Until that decision lands, FNV-1a is the algorithm
    every payload is hashed with.

    Reference: http://www.isthe.com/chongo/tech/comp/fnv/index.html.
    Known vectors (asserted in tests):
    - ``b"a"`` → ``e40c292c``
    - ``b"foobar"`` → ``bf9cf968``
    - ``b""`` → ``811c9dc5`` (offset basis)
    """

    @property
    def algorithm(self) -> str:
        return "fnv1a-32"

    def hash(self, payload: bytes) -> str:
        h = _FNV1A_32_OFFSET_BASIS
        for byte in payload:
            h ^= byte
            h = (h * _FNV1A_32_PRIME) & 0xFFFFFFFF
        return format(h, "08x")


# Register the default algorithm at import time so the module is
# usable without a bootstrap step. The decision on whether to keep
# FNV-1a or replace it lives in issue #659.
default_registry.register(FNV1A32())


__all__ = [
    "FNV1A32",
    "HashAlgorithm",
    "HashRegistry",
    "HashVersioned",
    "canonicalize",
    "compute_hash",
    "default_registry",
]
