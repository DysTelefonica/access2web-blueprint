# HARNESS-PROVENANCE: deterministic-quality-harness v1.6
# + architectural-guards-over-metrics DG-1, DG-2, DG-3, DG-4, DG-7, DG-13
# — test_decision_guards_wiring.py
"""Wiring pin for `scripts/check_decision_guards.py` (DG-1..DG-4, DG-7, DG-13).

Exercises the parsing half of the contract on synthetic markdown strings so
the scanner cannot be vacuously green (DG-2, DG-13) and the fail-closed
behaviour cannot drift away from the spec (DG-3).
"""

from __future__ import annotations

import importlib.util
import sys
import tempfile
from pathlib import Path

import pytest


def _find_script() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / "check_decision_guards.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_decision_guards.py not found")


@pytest.fixture(scope="module")
def script() -> Path:
    return _find_script()


@pytest.fixture(scope="module")
def gate_module(script: Path):
    spec = importlib.util.spec_from_file_location("check_decision_guards", script)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load spec for {script}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_decision_guards"] = module
    spec.loader.exec_module(module)
    return module


# Synthetic markdown helpers. Each fixture lives only in memory.

_DECISION_TABLE = (
    "| ID | Decision | Rationale |\n"
    "|---|---|---|\n"
    "| DG-1 | The decision registry is the markdown table. | Single source. |\n"
    "| DG-2 | A table is a decision table if first header cell normalises to id. | Robust. |\n"
    "| DG-3 | Fail-closed on missing input. | Hard Rule 18. |\n"
)

_WIRING_TABLE_WITH_LIST_CELLS = (
    "| Decisiones | Forma de cobertura | Artefacto |\n"
    "|---|---|---|\n"
    "| DG-1, DG-2, DG-3, DG-4 | guard | tests/lanzadera/test_decision_guards_wiring.py |\n"
    "| DG-11 | covered-by | scripts/check_complexity.py |\n"
)

_RISK_TABLE = (
    "| ID | Riesgo | Mitigacion | Gate |\n"
    "|---|---|---|---|\n"
    "| R-1 (H1) | Legacy hashing | Pin test | check_legacy_hashes.py |\n"
    "| R-2 (H2) | Drift | Verificacion runtime | make verify |\n"
)

_RANDOM_TWO_COLUMN_TABLE = (
    "| Recurso | Ruta |\n"
    "|---|---|\n"
    "| Decisiones D5-D91 | docs/08-decisiones-y-preguntas-abiertas.md |\n"
)


class _TempPath:
    """Context manager wrapping ``tempfile.TemporaryDirectory`` as a Path."""

    def __enter__(self) -> Path:
        self._tmp = tempfile.TemporaryDirectory()
        return Path(self._tmp.__enter__())

    def __exit__(self, exc_type, exc, tb) -> None:
        self._tmp.__exit__(exc_type, exc, tb)


pytest.TempPath = _TempPath  # type: ignore[attr-defined]


def _write_design_md(root: Path, relative: str, body: str) -> Path:
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(body, encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# DG-2 + DG-13: decision row recognition and table discrimination.
# ---------------------------------------------------------------------------


def test_parse_decision_tables_extracts_single_token_rows(gate_module) -> None:
    with pytest.TempPath() as tmp:
        _write_design_md(tmp, "openspec/changes/sample/design.md", _DECISION_TABLE)
        decisions = gate_module.parse_decision_tables(tmp)
    assert "DG-1" in decisions
    assert "DG-2" in decisions
    assert "DG-3" in decisions
    for _decision_id, (file_path, line) in decisions.items():
        assert file_path.endswith("design.md")
        assert line > 0


def test_parse_decision_tables_ignores_list_cells(gate_module) -> None:
    """DG-13: a cell listing ``DG-1, DG-2, DG-3`` is a wiring cell, not a
    decision row. The wiring table's header does NOT normalise to ``id``."""
    with pytest.TempPath() as tmp:
        _write_design_md(tmp, "openspec/changes/sample/design.md", _WIRING_TABLE_WITH_LIST_CELLS)
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}, f"wiring table must be skipped; got {sorted(decisions)}"


def test_parse_decision_tables_ignores_risk_table(gate_module) -> None:
    """DG-2: ``R-*`` rows are not decision rows because their prefix is outside
    ``DECISION_ID_PREFIXES``."""
    with pytest.TempPath() as tmp:
        _write_design_md(tmp, "openspec/changes/sample/design.md", _RISK_TABLE)
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}, f"risk table must not contribute; got {sorted(decisions)}"


def test_parse_decision_tables_ignores_non_id_header(gate_module) -> None:
    """DG-2: the ``Recurso | Ruta`` table has no header cell that normalises
    to ``id``; it must be skipped."""
    with pytest.TempPath() as tmp:
        _write_design_md(tmp, "openspec/changes/sample/design.md", _RANDOM_TWO_COLUMN_TABLE)
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}


# ---------------------------------------------------------------------------
# DG-3: fail-closed on malformed or absent input.
# ---------------------------------------------------------------------------


def test_parse_decision_tables_returns_empty_when_no_design_md(gate_module) -> None:
    with pytest.TempPath() as tmp:
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}


def test_parse_decision_tables_returns_empty_when_design_md_has_no_table(
    gate_module,
) -> None:
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "openspec/changes/sample/design.md",
            "# Design\n\nNo tables here, just prose.\n",
        )
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}


def test_parse_decision_tables_fails_closed_on_malformed_table(gate_module) -> None:
    """DG-3: a table whose header normalises to ``id`` but whose delimiter
    row is missing contributes zero decisions."""
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "openspec/changes/sample/design.md",
            ("| ID | Decision |\n| DG-1 | malformed |\n"),
        )
        decisions = gate_module.parse_decision_tables(tmp)
    assert decisions == {}


# ---------------------------------------------------------------------------
# DG-7: global uniqueness across the tree.
# ---------------------------------------------------------------------------


def test_parse_decision_tables_collects_global_ids_across_files(gate_module) -> None:
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "openspec/changes/alpha/design.md",
            "| ID | Decision |\n|---|---|\n| DA-1 | alpha decision |\n",
        )
        _write_design_md(
            tmp,
            "openspec/changes/beta/design.md",
            "| ID | Decision |\n|---|---|\n| DA-1 | beta decision |\n",
        )
        decisions = gate_module.parse_decision_tables(tmp)
    assert set(decisions) == {"DA-1"}


# ---------------------------------------------------------------------------
# Coverage layer (validate_coverage): DG-3, DG-7, DG-10.
# ---------------------------------------------------------------------------


def test_validate_coverage_emits_no_decision_table_on_empty_subject_set(
    gate_module,
) -> None:
    """DG-3 + Hard Rule 18: empty subject set fails closed."""
    with pytest.TempPath() as tmp:
        findings, _notes, _uncovered, _expired = gate_module.validate_coverage(
            tmp, decisions={}, claims={}, covered_by={}, baseline={}
        )
    keys = {finding["key"] for finding in findings}
    assert "no_decision_table" in keys


def test_validate_coverage_emits_duplicate_id_when_id_collides(gate_module) -> None:
    """DG-7: two ``design.md`` files contributing the same ID produce a
    ``duplicate_id`` finding."""
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "openspec/changes/alpha/design.md",
            "| ID | Decision |\n|---|---|\n| DA-1 | alpha |\n",
        )
        _write_design_md(
            tmp,
            "openspec/changes/beta/design.md",
            "| ID | Decision |\n|---|---|\n| DA-1 | beta |\n",
        )
        decisions = gate_module.parse_decision_tables(tmp)
        findings, _notes, _uncovered, _expired = gate_module.validate_coverage(
            tmp, decisions=decisions, claims={}, covered_by={}, baseline={}
        )
    keys = {finding["key"] for finding in findings}
    duplicate = [f for f in findings if f["key"] == "duplicate_id"]
    assert "duplicate_id" in keys
    assert duplicate[0]["id"] == "DA-1"


def test_validate_coverage_emits_uncovered_when_no_guard_or_baseline(
    gate_module,
) -> None:
    with pytest.TempPath() as tmp:
        _write_design_md(tmp, "openspec/changes/sample/design.md", _DECISION_TABLE)
        decisions = gate_module.parse_decision_tables(tmp)
        findings, _notes, uncovered, _expired = gate_module.validate_coverage(
            tmp, decisions=decisions, claims={}, covered_by={}, baseline={}
        )
    uncovered_keys = {f["key"] for f in findings if f["key"] == "uncovered"}
    assert uncovered_keys == {"uncovered"}
    assert {row["id"] for row in uncovered} == {"DG-1", "DG-2", "DG-3"}


def test_validate_coverage_accepts_guard_claim(gate_module) -> None:
    """A guard test whose ``HARNESS-PROVENANCE`` header cites the same ID
    counts as coverage."""
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "openspec/changes/sample/design.md",
            "| ID | Decision |\n|---|---|\n| DA-1 | sample |\n",
        )
        _write_design_md(
            tmp,
            "tests/sample/test_da1.py",
            (
                "# HARNESS-PROVENANCE: deterministic-quality-harness v1.6\n"
                "# + sample DA-1 — test_da1.py\n"
                '"""Pin for DA-1."""\n'
            ),
        )
        decisions = gate_module.parse_decision_tables(tmp)
        claims = gate_module.collect_guard_claims(tmp)
        findings, _notes, uncovered, _expired = gate_module.validate_coverage(
            tmp, decisions=decisions, claims=claims, covered_by={}, baseline={}
        )
    assert findings == []
    assert uncovered == []


def test_validate_coverage_emits_orphan_guard(gate_module) -> None:
    """DG-10: a guard header that cites an ID outside any decision table
    produces an ``orphan_guard`` finding."""
    with pytest.TempPath() as tmp:
        _write_design_md(
            tmp,
            "tests/sample/test_orphan.py",
            (
                "# HARNESS-PROVENANCE: deterministic-quality-harness v1.6\n"
                "# + sample DA-999 — test_orphan.py\n"
                '"""Guards a decision that does not exist."""\n'
            ),
        )
        claims = gate_module.collect_guard_claims(tmp)
        findings, _notes, _uncovered, _expired = gate_module.validate_coverage(
            tmp, decisions={}, claims=claims, covered_by={}, baseline={}
        )
    keys = {finding["key"] for finding in findings}
    assert "orphan_guard" in keys
    assert "no_decision_table" in keys
