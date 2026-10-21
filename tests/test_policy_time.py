"""The wall clock is an orchestration input, never hidden gate state (DQH 2.2)."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


def _load(name: str):
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / f"{name}.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def test_policy_date_parser_accepts_only_iso_calendar_dates() -> None:
    module = _load("policy_time")

    assert module.parse_policy_date("2026-08-12").isoformat() == "2026-08-12"
    for invalid in ("", "2026-8-12", "12/08/2026", "2026-02-30"):
        with pytest.raises(module.PolicyTimeError):
            module.parse_policy_date(invalid)
