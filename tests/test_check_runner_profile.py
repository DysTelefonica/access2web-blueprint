# HARNESS-PROVENANCE: adapted from ardelperal/APAP_WEB tests/test_check_required_jobs.py
"""Tests for the P51_runners_profile declaration and the operating-doc-drift
trigger surface (issue #858).

Issue #858 declares the runner profile in the consumer's `ci-pattern.yaml`
(`P51_runners_profile: hosted`) and removes the `push: branches: [main]`
trigger from `.github/workflows/operating-doc-drift.yml` so the drift gate
stays out of the PR path — it no longer hunts code defects, only doc drift
on `schedule` and `workflow_dispatch`.

These tests are structural: they read the on-disk YAML and assert the
declaration is present (or absent). A future operator who edits the yaml
without re-asserting the gate will trip one of them, which is the whole
reason this gate exists.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import yaml

# Make `scripts/` importable without sys.path games, same convention as
# tests/test_check_workflows.py and tests/test_check_required_jobs.py.
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

REPO_ROOT = Path(__file__).resolve().parent.parent
CI_PATTERN_YAML = REPO_ROOT / "ci-pattern.yaml"
OPERATING_DOC_DRIFT_YML = REPO_ROOT / ".github" / "workflows" / "operating-doc-drift.yml"


def _load_ci_pattern_yaml() -> dict[str, object]:
    return yaml.safe_load(CI_PATTERN_YAML.read_text(encoding="utf-8"))


def _load_operating_doc_drift() -> dict[str, object]:
    return yaml.safe_load(OPERATING_DOC_DRIFT_YML.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# P51 declaration in ci-pattern.yaml
# ---------------------------------------------------------------------------


def test_ci_pattern_yaml_declares_p51_runners_profile_hosted() -> None:
    """`P51_runners_profile: hosted` is the explicit declaration of HR-61.

    The gate `check_runner_profile.py` from the catalog is innocuous by
    default (retrocompatible), so the only way to assert the consumer's
    intent is to require the key to be present in the yaml — a comment
    does not survive a regenerate; a key does.
    """
    data = _load_ci_pattern_yaml()
    assert isinstance(data, dict)
    assert "P51_runners_profile" in data, (
        "ci-pattern.yaml must declare P51_runners_profile explicitly "
        "(HR-61). The catalog default is `hosted`, so the silent-default "
        "case hides the rule from the next operator."
    )
    assert data["P51_runners_profile"] == "hosted", (
        f"P51_runners_profile must be 'hosted' for this public repo "
        f"(HR-61), got {data['P51_runners_profile']!r}. A self-hosted "
        f"declaration requires P52_runners_pool_labels populated."
    )


# ---------------------------------------------------------------------------
# operating-doc-drift trigger surface
# ---------------------------------------------------------------------------


def _triggers(doc: dict[str, object]) -> object:
    """Return the value of `on:`, accepting PyYAML's `True` for plain `on`."""
    return doc.get("on", doc.get(True))


def test_operating_doc_drift_has_no_pull_request_trigger() -> None:
    """Criterion #2 of #858: no PR-reachable job depends on operating-doc-drift.

    The drift gate is doc-shape only; it does not catch code defects, so
    running it on a PR adds CI minutes for zero reviewer value.
    """
    doc = _load_operating_doc_drift()
    triggers = _triggers(doc)
    if isinstance(triggers, str):
        assert triggers != "pull_request", (
            "operating-doc-drift.yml must not run on pull_request "
            "(issue #858, criterion #2)."
        )
    elif isinstance(triggers, list):
        assert "pull_request" not in triggers, (
            "operating-doc-drift.yml must not list `pull_request` in its "
            "trigger list (issue #858, criterion #2)."
        )
    elif isinstance(triggers, dict):
        assert "pull_request" not in triggers, (
            "operating-doc-drift.yml must not declare a `pull_request` "
            "trigger (issue #858, criterion #2)."
        )


def test_operating_doc_drift_has_no_push_to_main_trigger() -> None:
    """Decision (issue #858): the drift gate moves out of the PR path.

    `push: branches: [main]` was the historical trigger; it kept the gate
    cheap on every push to the protected branch, but the gate runs against
    the SAME pin the consumer commits and offers no value-add over the
    PR-time check (the doc is regenerated locally before commit). The
    maintainer chose to keep the gate only on `schedule` and
    `workflow_dispatch`.
    """
    doc = _load_operating_doc_drift()
    triggers = _triggers(doc)
    assert not (isinstance(triggers, dict) and "push" in triggers), (
        "operating-doc-drift.yml must not declare a `push:` trigger "
        "(issue #858): the drift gate stays on `schedule` and "
        "`workflow_dispatch` only."
    )


def test_operating_doc_drift_keeps_schedule_and_workflow_dispatch() -> None:
    """The drift gate must still fire weekly + on manual dispatch.

    The schedule and the manual trigger are the only signals that catch
    catalog-pin drift over time (the consumer's `ci-pattern.yaml` is not
    exercised on every push anymore), so removing the trigger surface
    would silently drop the gate.
    """
    doc = _load_operating_doc_drift()
    triggers = _triggers(doc)
    assert isinstance(triggers, dict), (
        f"operating-doc-drift.yml must declare its triggers as a mapping "
        f"with at least `schedule` and `workflow_dispatch` keys; got "
        f"{triggers!r}"
    )
    assert "schedule" in triggers, (
        "operating-doc-drift.yml must keep the `schedule` trigger "
        "(issue #858): without it the catalog-pin drift is not detected."
    )
    assert "workflow_dispatch" in triggers, (
        "operating-doc-drift.yml must add the `workflow_dispatch` "
        "trigger (issue #858): manual re-runs are part of the gate's "
        "operating surface."
    )


# ---------------------------------------------------------------------------
# Pin sync: ci-pattern.yaml comment and the workflow's checkout ref
# ---------------------------------------------------------------------------


def _catalog_pin_from_yaml() -> str:
    """Extract the 40-hex SHA that ci-pattern.yaml names in its header comment."""
    text = CI_PATTERN_YAML.read_text(encoding="utf-8")
    # The header is a YAML comment block: lines starting with `#`. The pin
    # always appears in a single line as a bare 40-hex SHA.
    candidates = re.findall(r"^#.*?\b([0-9a-f]{40})\b", text, flags=re.MULTILINE)
    assert candidates, (
        "ci-pattern.yaml header must name the catalog pin as a 40-hex SHA "
        "in a comment line"
    )
    return candidates[0]


def _catalog_pin_from_workflow() -> str:
    """Extract the `ref:` of the `actions/checkout` that pulls DysTelefonica/team-skills."""
    doc = yaml.safe_load(OPERATING_DOC_DRIFT_YML.read_text(encoding="utf-8"))
    jobs = doc.get("jobs") or {}
    for job in jobs.values():
        if not isinstance(job, dict):
            continue
        for step in job.get("steps") or []:
            if not isinstance(step, dict):
                continue
            uses = str(step.get("uses") or "")
            with_block = step.get("with") or {}
            if uses.startswith("actions/checkout@") and with_block.get("repository") == (
                "DysTelefonica/team-skills"
            ):
                ref = with_block.get("ref")
                assert isinstance(ref, str) and re.fullmatch(
                    r"[0-9a-f]{40}", ref
                ), (
                    "operating-doc-drift.yml must check out the catalog "
                    f"pin as a 40-hex SHA, got ref={ref!r}"
                )
                return ref
    raise AssertionError(
        "operating-doc-drift.yml must check out DysTelefonica/team-skills "
        "with a 40-hex SHA on a step that uses actions/checkout"
    )


def test_ci_pattern_yaml_pin_comment_matches_operating_doc_drift_ref() -> None:
    """The catalog pin in the yaml comment and the checkout ref in the
    workflow must agree.

    The yaml's header comment names the pin for human readers; the
    workflow's `actions/checkout` uses it for the drift measurement. If
    they drift, the comment lies and the next operator will regenerate
    the doc against a pin the drift gate does not actually use.
    """
    yaml_pin = _catalog_pin_from_yaml()
    workflow_pin = _catalog_pin_from_workflow()
    assert yaml_pin == workflow_pin, (
        f"ci-pattern.yaml names pin {yaml_pin} but "
        f"operating-doc-drift.yml checks out {workflow_pin}; the two "
        f"must agree so the comment and the gate measure against the "
        f"same catalog commit."
    )
