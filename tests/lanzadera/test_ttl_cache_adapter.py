"""Contract-conformance test for :class:`TtlCacheAdapter`.

DA-8, D70, D71, D74. The adapter is the in-process TTL cache; the
candidates are slow-changing derived reads (``apps.list_active``,
``profiles.list_for_app``, the effective-permissions map).

The MVP backing is ``cachetools.TTLCache``, which only supports a
cache-wide TTL at construction. The per-call ``set(key, value, ttl)``
in earlier revisions of the WU is gone because the underlying
``TTLCache.__setitem__`` cannot honour it without wrapping the value
(which would break ``get``). The composition root picks the right
``default_ttl_seconds`` at adapter construction (D70).
"""

from __future__ import annotations

import time

import pytest

from app.src.modules.lanzadera.adapters.cross.cache import TtlCacheAdapter
from app.src.modules.lanzadera.ports.cache import CachePort


def test_ttl_cache_satisfies_protocol() -> None:
    """Static: the adapter implements the Protocol."""
    adapter: CachePort = TtlCacheAdapter()
    assert hasattr(adapter, "get")
    assert hasattr(adapter, "set")
    assert hasattr(adapter, "invalidate")


def test_ttl_cache_get_returns_none_for_absent_key() -> None:
    adapter = TtlCacheAdapter()
    assert adapter.get("never-set") is None


def test_ttl_cache_set_then_get_roundtrips() -> None:
    adapter = TtlCacheAdapter()
    adapter.set("k1", "v1")
    assert adapter.get("k1") == "v1"


def test_ttl_cache_supports_arbitrary_value_types() -> None:
    """The port contract is ``Any`` — strings, lists, dicts, tuples all pass through."""
    adapter = TtlCacheAdapter()
    for key, value in (
        ("str", "hello"),
        ("list", [1, 2, 3]),
        ("dict", {"a": 1}),
        ("none", None),  # ``None`` is a valid cached value, distinct from absent
        ("int", 42),
    ):
        adapter.set(key, value)
        assert adapter.get(key) == value


def test_ttl_cache_overwrites_existing_value() -> None:
    """Same key, second ``set`` replaces the value (last-write-wins)."""
    adapter = TtlCacheAdapter()
    adapter.set("k", "v1")
    adapter.set("k", "v2")
    assert adapter.get("k") == "v2"


def test_ttl_cache_respects_ttl_and_expires() -> None:
    """An entry expires after the configured TTL elapses (DA-8)."""
    adapter = TtlCacheAdapter(default_ttl_seconds=1)
    adapter.set("short-lived", "v")
    assert adapter.get("short-lived") == "v"
    time.sleep(1.2)  # TTL + a small margin
    assert adapter.get("short-lived") is None


def test_ttl_cache_rejects_non_positive_default_ttl() -> None:
    """D70: a non-expiring (zero or negative TTL) entry would hide invalidation bugs."""
    with pytest.raises(ValueError, match="D70"):
        TtlCacheAdapter(default_ttl_seconds=0)
    with pytest.raises(ValueError, match="D70"):
        TtlCacheAdapter(default_ttl_seconds=-1)


def test_ttl_cache_invalidate_drops_matching_keys() -> None:
    """``invalidate(prefix)`` drops every entry whose key starts with ``prefix``."""
    adapter = TtlCacheAdapter()
    adapter.set("apps:list", [1, 2])
    adapter.set("apps:active", [3])
    adapter.set("profiles:42", {"code": "default"})
    adapter.set("user:1:perms", ["read"])

    dropped = adapter.invalidate("apps:")
    assert dropped == 2
    assert adapter.get("apps:list") is None
    assert adapter.get("apps:active") is None
    assert adapter.get("profiles:42") == {"code": "default"}
    assert adapter.get("user:1:perms") == ["read"]


def test_ttl_cache_invalidate_returns_count() -> None:
    adapter = TtlCacheAdapter()
    assert adapter.invalidate("anything") == 0
    adapter.set("a:1", 1)
    adapter.set("a:2", 2)
    adapter.set("b:1", 3)
    assert adapter.invalidate("a:") == 2
    assert adapter.invalidate("a:") == 0  # second call: nothing matches
    assert adapter.get("b:1") == 3


def test_ttl_cache_invalidate_empty_prefix_drops_everything() -> None:
    """An empty prefix matches every key — useful for full-flush scenarios."""
    adapter = TtlCacheAdapter()
    adapter.set("k1", 1)
    adapter.set("k2", 2)
    dropped = adapter.invalidate("")
    assert dropped == 2
    assert adapter.get("k1") is None
    assert adapter.get("k2") is None


def test_ttl_cache_invalidate_no_match() -> None:
    adapter = TtlCacheAdapter()
    adapter.set("k1", 1)
    assert adapter.invalidate("nope:") == 0
    assert adapter.get("k1") == 1  # not touched
