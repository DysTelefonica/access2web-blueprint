#!/usr/bin/env python3
# ci-pattern asset — changelog consolidation for the release PR (HR-44, DysTelefonica/team-skills#665)
"""Consolidate pending changelog fragments into a release of one skill.

Run by hand inside the release PR, never by CI and never on main (main is
PR-only). For the skill directory it:

1. reads every fragment in ``<skill>/changelog.d/*.md`` (``README.md`` excluded);
2. validates each one: first line ``type: patch`` or ``type: minor``, then an entry;
3. takes the highest type (``minor`` if any fragment is minor, else ``patch``)
   and computes the next version from ``metadata.version`` in SKILL.md;
4. writes that version into SKILL.md, adds one bullet (one sub-bullet per fragment) under the first
   ``### Changed`` of ``## [Unreleased]`` in CHANGELOG.md, and deletes the
   consumed fragments.

Without ``--write`` it only prints the plan. With no pending fragments it
changes nothing and exits 0.

Fail-closed (HR-3): an invalid fragment, an unreadable or missing version,
or a CHANGELOG without ``## [Unreleased]`` / ``### Changed`` exits 2 and
writes nothing.

Exit codes: 0 done or nothing pending · 1 plan printed without --write
(pending fragments exist) · 2 invalid input.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

FRAGMENT_DIR = "changelog.d"
TYPE_LINE = re.compile(r"type:\s*(patch|minor)")
VERSION = re.compile(r"^(\s*version:\s*[\"']?)(\d+)\.(\d+)\.(\d+)([\"']?\s*)$", re.M)


class ConsolidateError(Exception):
    """A condition the consolidation refuses to interpret; maps to exit 2."""


def read_fragments(skill: Path) -> list[tuple[Path, str, str]]:
    """Return (path, type, entry) for every pending fragment, sorted by name."""
    out = []
    for path in sorted((skill / FRAGMENT_DIR).glob("*.md")):
        if path.name == "README.md":
            continue
        lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()
                 if line.strip()]
        if not lines or not TYPE_LINE.fullmatch(lines[0]):
            raise ConsolidateError(
                f"{path}: first line must be 'type: patch' or 'type: minor'")
        if len(lines) < 2:
            raise ConsolidateError(f"{path}: fragment has no entry after the type line")
        out.append((path, TYPE_LINE.fullmatch(lines[0]).group(1), " ".join(lines[1:])))
    return out


def next_version(skill_md: Path, kind: str) -> tuple[str, str, str]:
    """Return (old, new, new_skill_md_text) from metadata.version."""
    text = skill_md.read_text(encoding="utf-8")
    m = VERSION.search(text)
    if not m:
        raise ConsolidateError(f"{skill_md}: metadata.version not found")
    major, minor, patch = int(m.group(2)), int(m.group(3)), int(m.group(4))
    old = f"{major}.{minor}.{patch}"
    if kind == "minor":
        new = f"{major}.{minor + 1}.0"
    else:
        new = f"{major}.{minor}.{patch + 1}"
    new_text = VERSION.sub(lambda mm: f"{mm.group(1)}{new}{mm.group(5)}", text, count=1)
    return old, new, new_text


def add_changelog_bullet(changelog: Path, skill_name: str, old: str, new: str,
                         entries: list[str]) -> str:
    text = changelog.read_text(encoding="utf-8")
    unreleased = text.find("## [Unreleased]")
    if unreleased < 0:
        raise ConsolidateError(f"{changelog}: '## [Unreleased]' not found")
    changed = text.find("### Changed", unreleased)
    if changed < 0:
        raise ConsolidateError(f"{changelog}: '### Changed' not found under [Unreleased]")
    insert_at = text.index("\n", changed) + 1
    # One sub-bullet per fragment under the version header, so a release with
    # many fragments stays readable instead of one huge paragraph.
    bullet = (f"- **`{skill_name}`** (v{old} → v{new}):\n"
              + "".join(f"  - {entry}\n" for entry in entries))
    return text[:insert_at] + "\n" + bullet + text[insert_at:]


def consolidate(skill: Path, changelog: Path, write: bool) -> int:
    fragments = read_fragments(skill)
    if not fragments:
        print(f"CONSOLIDATE: nada pendiente en {skill / FRAGMENT_DIR}")
        return 0
    kind = "minor" if any(t == "minor" for _, t, _ in fragments) else "patch"
    old, new, new_skill_text = next_version(skill / "SKILL.md", kind)
    entries = [entry for _, _, entry in fragments]
    new_changelog = add_changelog_bullet(changelog, skill.name, old, new, entries)
    print(f"CONSOLIDATE: {len(fragments)} fragmento(s), tipo {kind}: v{old} → v{new}")
    for path, _, _ in fragments:
        print(f"  - consume {path.name}")
    if not write:
        print("CONSOLIDATE: plan only; rerun with --write to apply")
        return 1
    (skill / "SKILL.md").write_text(new_skill_text, encoding="utf-8")
    changelog.write_text(new_changelog, encoding="utf-8")
    for path, _, _ in fragments:
        path.unlink()
    return 0


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv):
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Consolidate changelog fragments (HR-44).")
    ap.add_argument("--skill-dir", required=True)
    ap.add_argument("--changelog", default="CHANGELOG.md")
    ap.add_argument("--write", action="store_true")
    args = ap.parse_args(argv)
    try:
        return consolidate(Path(args.skill_dir), Path(args.changelog), args.write)
    except (ConsolidateError, OSError) as exc:
        print(f"CONSOLIDATE ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
