# HARNESS-PROVENANCE: deterministic-quality-harness v2.3 — assets/tests/test_cache_consistency.py
"""A derived cache is checked against its source derivation for every live row (Hard Rule 20)."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import MappingProxyType

ASSETS = Path(__file__).resolve().parent.parent
SCRIPTS = ASSETS / "scripts"


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / f"{name}.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


SOURCE = {"a1": "shelter", "a2": "foster", "a3": "adopted"}


def _derive(key: str) -> str:
    return SOURCE[key]


def test_consistent_cache_is_green() -> None:
    module = _load("cache_consistency")
    cached = MappingProxyType(dict(SOURCE))

    drift = module.compare(sorted(SOURCE), cached, _derive)

    assert drift == []
    assert module.exit_code(drift) == module.EXIT_OK


def test_check_goes_red_on_a_deliberately_stale_row() -> None:
    module = _load("cache_consistency")
    cached = MappingProxyType({**SOURCE, "a2": "shelter"})

    drift = module.compare(sorted(SOURCE), cached, _derive)

    assert drift == [module.Drift(key="a2", cached="shelter", derived="foster")]
    assert drift[0].kind == "mismatch"
    assert module.exit_code(drift) == module.EXIT_DRIFT
    assert "a2" in module.summary(drift, sample_size=5)


def test_live_row_missing_from_the_cache_is_drift() -> None:
    module = _load("cache_consistency")
    cached = MappingProxyType({"a1": "shelter", "a3": "adopted"})

    drift = module.compare(sorted(SOURCE), cached, _derive)

    assert [(d.key, d.kind) for d in drift] == [("a2", "missing")]


def test_summary_counts_every_row_but_samples_only_a_few() -> None:
    module = _load("cache_consistency")
    drift = [module.Drift(key=f"k{i}", cached="x", derived="y") for i in range(30)]

    text = module.summary(drift, sample_size=3)

    assert text.splitlines()[0] == "differences: 30"
    assert "k2" in text
    assert "k3" not in text
