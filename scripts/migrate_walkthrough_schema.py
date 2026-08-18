#!/usr/bin/env python3
"""HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — migrate_walkthrough_schema.py.

One-shot migration that brings the 43 real walkthroughs under the v2.1 schema
documented at `skills/documentation-alan-style/references/templates/walkthrough.json.tmpl`.

Adds (only when missing) or translates:
  * per-form `formName` from `form_name`/`name`/etc.
  * per-form `sourcePath` from `source_path`.
  * per-form `tool_warnings: []`.
  * per-form `codegraph_summary` (inherited for sub-forms).
  * per-form `ui`/`geometry`/`behavior` built from flat fields.
  * per-form `verify_form_bindings` from `verify_form_bindings_status`; backfill status.
  * per-form `unattended` from `unattended_detection`.
  * top-level `summary` derived from `forms[]` when missing.

This script is idempotent: running it twice produces no further changes.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

WALKTHROUGH_GLOB = "docs/03-aplicaciones/*/walkthrough-*.json"
SUBFORM_PATTERN = re.compile(r"(?i)subfrm|subform|sub_")


def _git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], capture_output=True, text=True, check=False, encoding="utf-8"
    )
    if result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def discover_walkthroughs(repo_root: Path) -> list[Path]:
    raw = _git("-C", str(repo_root), "ls-files", WALKTHROUGH_GLOB)
    return sorted(repo_root / line.strip() for line in raw.splitlines() if line.strip())


def derive_summary(forms: list, data: dict) -> dict:
    attended = sum(1 for f in forms if isinstance(f, dict) and f.get("unattended") is False)
    unattended = sum(1 for f in forms if isinstance(f, dict) and f.get("unattended") is True)
    return {
        "total_forms": len(forms),
        "attended_forms": attended,
        "unattended_forms": unattended,
        "known_skipped_tools": data.get("known_skipped_tools", []),
        "known_degraded_tools": data.get("known_degraded_tools", []),
    }


def is_subform(form: dict) -> bool:
    name = form.get("formName") or form.get("form_name") or form.get("name") or ""
    return bool(SUBFORM_PATTERN.search(name))


def default_codegraph_summary(form: dict) -> dict:
    if is_subform(form):
        return {
            "status": "inherited",
            "note": "Sub-form; shares parent codegraph context.",
            "stale_banner_ignored": True,
        }
    return {
        "status": "not_applicable",
        "note": "No codegraph evidence collected for this form.",
        "stale_banner_ignored": True,
    }


def default_verify_form_bindings(form: dict) -> dict:
    status = form.get("verify_form_bindings_status") or "skipped_tool_broken"
    error = (
        form.get("verify_form_bindings_error")
        or "verify_form_bindings returns RESULT_CONTRACT_VIOLATION (issue #1412)."
    )
    return {"status": status, "reason": error, "findings": []}


def build_ui_from_flat(form: dict) -> dict:
    return {
        "controls_count": form.get("controls_count", 0),
        "formEvents": form.get("form_events", []),
        "bindings_count": form.get("bindings_count", 0),
        "warnings": [],
    }


def build_behavior_from_flat(form: dict) -> dict:
    return {
        "controls_count": form.get("controls_count", 0),
        "controls_with_handler_evidence": form.get("event_handlers_declared", 0),
        "tables_referenced": [],
        "codegraph_evidence_status": form.get("codegraph_evidence_status", "not_applicable"),
    }


def build_geometry_from_flat(form: dict) -> dict:
    findings = form.get("layout_findings", [])
    return {
        "controls_count": form.get("controls_count", 0),
        "findings_count": len(findings) if isinstance(findings, list) else 0,
        "status": "ok" if not findings else "with_findings",
    }


def migrate_form(form: dict, changes: list[str]) -> None:
    if not isinstance(form, dict):
        return
    label = form.get("formName") or form.get("form_name") or form.get("name") or "<unnamed>"

    if "formName" not in form:
        for src in ("form_name", "name", "formName_in_ir", "form_class_name", "form"):
            if src in form and form[src]:
                form["formName"] = form[src]
                changes.append(f"{label}.formName<{src}")
                break
    if "sourcePath" not in form:
        for src in ("source_path",):
            if src in form and form[src]:
                form["sourcePath"] = form[src]
                changes.append(f"{label}.sourcePath<{src}")
                break

    if "tool_warnings" not in form:
        form["tool_warnings"] = []
        changes.append(f"{label}.tool_warnings=[]")

    if "codegraph_summary" not in form:
        form["codegraph_summary"] = default_codegraph_summary(form)
        kind = "inherited" if is_subform(form) else "not_applicable"
        changes.append(f"{label}.codegraph_summary={kind}")

    if "ui" not in form:
        form["ui"] = build_ui_from_flat(form)
        changes.append(f"{label}.ui=built_from_flat")

    if "geometry" not in form:
        form["geometry"] = build_geometry_from_flat(form)
        changes.append(f"{label}.geometry=built_from_flat")

    if "behavior" not in form:
        form["behavior"] = build_behavior_from_flat(form)
        changes.append(f"{label}.behavior=built_from_flat")

    if "verify_form_bindings" not in form:
        form["verify_form_bindings"] = default_verify_form_bindings(form)
        changes.append(f"{label}.verify_form_bindings=built_from_flat")
    else:
        vfb = form.get("verify_form_bindings")
        if isinstance(vfb, dict) and "status" not in vfb:
            vfb["status"] = vfb.get("verify_form_bindings_status") or "skipped_tool_broken"
            if "reason" not in vfb:
                vfb["reason"] = (
                    vfb.get("verify_form_bindings_error")
                    or "verify_form_bindings returns RESULT_CONTRACT_VIOLATION (issue #1412)."
                )
            changes.append(f"{label}.verify_form_bindings.status=backfilled")

    if "unattended" not in form:
        detection = form.get("unattended_detection")
        if detection is True:
            form["unattended"] = True
        elif detection is False:
            form["unattended"] = False
        else:
            form["unattended"] = False
        changes.append(f"{label}.unattended={form['unattended']}")


def migrate_file(path: Path) -> tuple[bool, list[str]]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return False, [f"PARSE_ERROR in {path}"]

    changes: list[str] = []
    forms = data.get("forms")
    if isinstance(forms, list):
        for form in forms:
            migrate_form(form, changes)

    if "summary" not in data and isinstance(forms, list):
        data["summary"] = derive_summary(forms, data)
        changes.append("summary={derived}")

    if changes:
        path.write_text(
            json.dumps(data, indent=2, ensure_ascii=False, sort_keys=False) + "\n",
            encoding="utf-8",
        )
    return bool(changes), changes


def _pin_output_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--repo-root",
        default=os.environ.get("REPO_ROOT", "."),
        help="path to repo root (default: $REPO_ROOT or cwd)",
    )
    args = parser.parse_args(argv)

    repo_root = Path(args.repo_root).resolve()
    walkthroughs = discover_walkthroughs(repo_root)

    total_changed = 0
    total_changes = 0
    for path in walkthroughs:
        changed, changes = migrate_file(path)
        if changed:
            total_changed += 1
            total_changes += len(changes)
            print(f"  {path.relative_to(repo_root)}: {len(changes)} change(s)")

    print(
        f"\nMigrated {total_changed} of {len(walkthroughs)} walkthrough(s); "
        f"{total_changes} field additions."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
