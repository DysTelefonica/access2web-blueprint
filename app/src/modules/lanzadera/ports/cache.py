"""CachePort — DA-8, D70, D71, D74.

In-process cache backed by ``cachetools.TTLCache`` for the MVP. The
candidates are catalogued in design.md: ``apps.list_active``,
``profiles.list_for_app``, and the effective-permissions map. D70
prohibits caching counters, volatile metrics, or session data — only
slow-changing, derived reads are eligible.

The TTL is per-entry (the caller picks, the adapter respects) and
invalidation is explicit (the caller fires ``invalidate(prefix)`` on
every mutation). Production swaps in Redis behind the same port
without touching the domain (D71).
"""

from __future__ import annotations

from typing import Any, Protocol


class CachePort(Protocol):
    """Driven port for the in-process TTL cache.

    The TTL is fixed at adapter construction (``default_ttl_seconds`` on
    the concrete adapter) — the MVP backing ``cachetools.TTLCache`` does
    not support per-entry TTL at ``__setitem__`` time, and a per-entry
    TTL that contradicts the cache-wide TTL would hide invalidation
    bugs. Callers that need a different TTL per entry pick the adapter
    at composition-root time (DA-8, D70).
    """

    def get(self, key: str) -> Any | None:
        """Return the cached value for ``key`` or ``None`` if absent / expired."""
        ...

    def set(self, key: str, value: Any) -> None:
        """Cache ``value`` under ``key`` for the adapter's default TTL.

        The TTL is fixed at adapter construction; the per-call
        ``ttl_seconds`` argument that earlier revisions carried is gone
        because the underlying ``cachetools.TTLCache`` cannot honour it
        without wrapping the value (which would break ``get``). For
        non-default TTL, configure the adapter with a different
        ``default_ttl_seconds`` at composition-root time.
        """
        ...

    def invalidate(self, prefix: str) -> int:
        """Drop every entry whose key starts with ``prefix``.

        Returns the number of entries dropped (for the operator
        harness and the application-layer observability).
        """
        ...
