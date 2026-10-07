"""Pins the governance contracts installed by the ci-pattern adoption, phase 3.

The canonical issue form and the pull-request template are DATA: the pattern's gates read
the six section names and the chain block out of them. A silent edit that renames a section
or drops a chain field would break those gates without any code changing, so the shape of
both files is asserted here.
"""

from __future__ import annotations

from pathlib import Path

import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
ISSUE_DIR = REPO_ROOT / ".github" / "ISSUE_TEMPLATE"
PR_TEMPLATE = REPO_ROOT / ".github" / "PULL_REQUEST_TEMPLATE.md"

#: Names and order fixed by ci-pattern (SKILL.md, «Operación por unidad de trabajo»).
CANONICAL_SECTIONS = (
    "Problema y contexto",
    "Evidencia verificable",
    "Alcance y no objetivos",
    "Criterios de aceptación",
    "Plan de validación",
    "Dependencias y riesgos",
)

#: Fields of the Chain Context block the PR-contract gate reads (HR-54).
CHAIN_FIELDS = (
    "chain",
    "position",
    "base",
    "depends-on",
    "follow-up",
    "starts-at",
    "ends-with",
    "review-budget",
)


def test_issue_form_declares_the_six_canonical_sections_in_order() -> None:
    body = yaml.safe_load((ISSUE_DIR / "issue-canonical.yml").read_text(encoding="utf-8"))["body"]
    fields = [block for block in body if block["type"] == "textarea"]

    assert [block["attributes"]["label"] for block in fields] == list(CANONICAL_SECTIONS)
    assert all(block["validations"]["required"] for block in fields)


def test_issue_config_disables_blank_issues() -> None:
    config = yaml.safe_load((ISSUE_DIR / "config.yml").read_text(encoding="utf-8"))

    assert config["blank_issues_enabled"] is False


def test_pull_request_template_declares_the_gate_contract() -> None:
    text = PR_TEMPLATE.read_text(encoding="utf-8")

    assert "## Chain Context" in text
    assert all(f"- {field}:" in text for field in CHAIN_FIELDS)
    assert text.count("📍") == 1
    assert "## Tests que prueban el cierre" in text
