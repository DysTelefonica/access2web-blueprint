"""``TtlCacheAdapter`` — DA-8, D70, D71, D74.

In-process cache backed by ``cachetools.TTLCache(maxsize=1024)``. The
candidates are slow-changing derived reads: ``apps.list_active``,
``profiles.list_for_app``, the effective-permissions map. D70
prohibits caching counters, volatile metrics, or session data — the
adapter exposes a generic ``get/set/invalidate(prefix)`` surface; the
caller is responsible for not storing the wrong things.

``invalidate(prefix)`` is the only way to drop entries today
(cachetools does not expose wildcard-delete). The prefix-match is
``str.startswith``; the caller picks a key shape that supports it
(every key in a category shares a prefix like ``\"apps:\"``).
"""

from __future__ import annotations

from typing import Any

from cachetools import TTLCache

from app.src.modules.lanzadera.ports.cache import CachePort


class TtlCacheAdapter(CachePort):
    """In-process TTL cache. One ``TTLCache`` instance per adapter.

    ``cachetools.TTLCache`` does not support per-entry TTL at
    ``__setitem__`` time (the API is ``__setitem__(key, value)`` only);
    the TTL is fixed at the cache-wide level. The composition root
    picks the right ``default_ttl_seconds`` for the deployment
    (DA-8, D70). D70 forbids non-expiring entries — a positive
    ``default_ttl_seconds`` is enforced at construction.
    """

    def __init__(self, maxsize: int = 1024, default_ttl_seconds: int = 60) -> None:
        if default_ttl_seconds <= 0:
            raise ValueError(
                f"default_ttl_seconds must be positive; got {default_ttl_seconds}. "
                "D70 forbids non-expiring entries."
            )
        self._cache: TTLCache[str, Any] = TTLCache(maxsize=maxsize, ttl=default_ttl_seconds)

    def get(self, key: str) -> Any | None:
        """Return the cached value for ``key`` or ``None`` if absent / expired."""
        try:
            return self._cache[key]
        except KeyError:
            return None

    def set(self, key: str, value: Any) -> None:
        """Cache ``value`` under ``key`` for the adapter's default TTL."""
        self._cache[key] = value

    def invalidate(self, prefix: str) -> int:
        """Drop every entry whose key starts with ``prefix``.

        Returns the number of entries dropped.
        """
        keys_to_drop = [k for k in self._cache if k.startswith(prefix)]
        for k in keys_to_drop:
            del self._cache[k]
        return len(keys_to_drop)
