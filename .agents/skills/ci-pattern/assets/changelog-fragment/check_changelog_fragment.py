#!/usr/bin/env python3
# ci-pattern asset — changelog fragment gate (HR-44, DysTelefonica/team-skills#665)
"""Changelog fragment gate: every material skill edit ships a fragment (HR-44).

Takes a base ref and a head ref and gates every skill directory touched by
the diff (a SKILL.md at head). Such a directory must gain at least one NEW
fragment under ``<skill>/changelog.d/*.md`` in the same diff. A fragment is
valid when its first non-empty line is ``type: patch`` or ``type: minor``
and a non-empty entry follows it. ``README.md`` inside ``changelog.d/`` is
documentation, not a fragment. Only fragments ADDED by the diff count: an
edit to a fragment already on base does not satisfy the rule.

No PR writes ``metadata.version`` or ``CHANGELOG.md`` for its own change:
the release PR consolidates pending fragments (issue #665).

The release PR itself is the one exception, recognised only by its exact
shape for a skill: it deletes at least one fragment, adds none, changes in
``SKILL.md`` only the ``metadata.version`` line, changes nothing else inside
the skill but a ``CHANGELOG.md``, and the diff touches a ``CHANGELOG.md``
(root or in the skill). Any other change in that skill keeps requiring a
fragment.

Fail-closed: a base ref that does not resolve, an empty merge-base, or a
fragment that cannot be read at head exits 2 (HR-3). The diff is taken
against the merge-base, so only commits the PR introduces are counted.

Exit codes: 0 clean · 1 findings · 2 invalid input.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

FRAGMENT_DIR = "changelog.d"
TYPE_LINE = re.compile(r"type:\s*(patch|minor)")
VERSION_LINE = re.compile(r"\s*version:\s*[\"']?\d+\.\d+\.\d+[\"']?\s*")


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def git(repo, *args) -> str:
    r = subprocess.run(["git", "-C", str(repo), *args],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise GateError(f"git {' '.join(args)} failed: {r.stderr.strip()}")
    return r.stdout.strip()


def rev_parse(repo: Path, ref: str) -> str:
    r = subprocess.run(["git", "-C", str(repo), "rev-parse", "--verify", "-q", ref],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise GateError(f"ref {ref!r} does not resolve")
    return r.stdout.strip()


def merge_base(repo: Path, base: str, head: str) -> str:
    out = git(repo, "merge-base", base, head)
    if not out:
        raise GateError(
            f"merge-base of {base!r} and {head!r} is empty: branches share "
            "no common ancestor")
    return out


def skill_exists(repo: Path, ref: str, skill_dir: str) -> bool:
    r = subprocess.run(["git", "-C", str(repo), "cat-file", "-e",
                        f"{ref}:{skill_dir}/SKILL.md"], capture_output=True)
    return r.returncode == 0


def discover_skill_dirs(repo: Path, mb: str, head: str) -> list[str]:
    dirs = set()
    for f in git(repo, "diff", "--name-only", f"{mb}..{head}").splitlines():
        parts = f.strip().split("/")
        for i in range(1, len(parts)):
            d = "/".join(parts[:i])
            if skill_exists(repo, head, d):
                dirs.add(d)
    return sorted(dirs)


def changed_files(repo: Path, mb: str, head: str, skill_dir: str) -> list[str]:
    out = git(repo, "diff", "--name-only", f"{mb}..{head}", "--", skill_dir)
    return [line for line in out.splitlines() if line.strip()]


def added_fragments(repo: Path, mb: str, head: str, skill_dir: str) -> list[str]:
    """Fragment paths ADDED by the diff (not modified, not pre-existing)."""
    out = git(repo, "diff", "--name-only", "--diff-filter=A",
              f"{mb}..{head}", "--", f"{skill_dir}/{FRAGMENT_DIR}")
    return [p for p in out.splitlines()
            if p.endswith(".md") and Path(p).name != "README.md"]


def deleted_fragments(repo: Path, mb: str, head: str, skill_dir: str) -> list[str]:
    out = git(repo, "diff", "--name-only", "--no-renames", "--diff-filter=D",
              f"{mb}..{head}", "--", f"{skill_dir}/{FRAGMENT_DIR}")
    return [p for p in out.splitlines()
            if p.endswith(".md") and Path(p).name != "README.md"]


def only_version_line_changed(repo: Path, mb: str, head: str, skill_dir: str) -> bool:
    """True when SKILL.md differs from the merge-base by exactly one
    ``metadata.version`` line (one line removed, one added)."""
    out = git(repo, "diff", "-U0", "--no-renames", f"{mb}..{head}", "--",
              f"{skill_dir}/SKILL.md")
    gone = [ln[1:] for ln in out.splitlines() if ln.startswith("-") and not ln.startswith("---")]
    came = [ln[1:] for ln in out.splitlines() if ln.startswith("+") and not ln.startswith("+++")]
    return (len(gone) == 1 and len(came) == 1
            and bool(VERSION_LINE.fullmatch(gone[0]))
            and bool(VERSION_LINE.fullmatch(came[0])) and gone[0] != came[0])


def is_release_shape(repo: Path, mb: str, head: str, skill_dir: str) -> bool:
    """The consolidation PR (HR-44), by its exact shape; see the module docstring."""
    if not deleted_fragments(repo, mb, head, skill_dir) \
            or added_fragments(repo, mb, head, skill_dir):
        return False
    if not only_version_line_changed(repo, mb, head, skill_dir):
        return False
    allowed_changelog = f"{skill_dir}/CHANGELOG.md"
    for f in changed_files(repo, mb, head, skill_dir):
        if f == f"{skill_dir}/SKILL.md" or f == allowed_changelog \
                or (f.startswith(f"{skill_dir}/{FRAGMENT_DIR}/") and f.endswith(".md")
                    and Path(f).name != "README.md"):
            continue
        return False
    touched = git(repo, "diff", "--name-only", "--no-renames", f"{mb}..{head}",
                  "--", "CHANGELOG.md", allowed_changelog)
    return bool(touched.strip())


def validate_fragment(text: str) -> str | None:
    """Return a finding reason, or None when the fragment is valid."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "empty fragment"
    if not TYPE_LINE.fullmatch(lines[0]):
        return (f"first line must be 'type: patch' or 'type: minor', "
                f"got {lines[0]!r}")
    if len(lines) < 2:
        return "fragment has no entry after the type line"
    return None


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv):
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Changelog fragment gate (HR-44).")
    ap.add_argument("--repo", required=True)
    ap.add_argument("--base", required=True)
    ap.add_argument("--head", required=True)
    args = ap.parse_args(argv)
    repo = Path(args.repo)
    report = []
    releases = []
    try:
        base_sha = rev_parse(repo, args.base)
        head_sha = rev_parse(repo, args.head)
        mb = merge_base(repo, base_sha, head_sha)
        for sd in discover_skill_dirs(repo, mb, head_sha):
            if not changed_files(repo, mb, head_sha, sd):
                continue
            if not skill_exists(repo, base_sha, sd):
                continue  # new skill: no base to edit, no fragment owed
            if is_release_shape(repo, mb, head_sha, sd):
                releases.append(sd)
                continue
            frags = added_fragments(repo, mb, head_sha, sd)
            if not frags:
                report.append(f"HR-44 skill {sd!r} changed without an added "
                              f"fragment in {sd}/{FRAGMENT_DIR}/*.md")
                continue
            for path in frags:
                try:
                    text = git(repo, "show", f"{head_sha}:{path}")
                except GateError as exc:
                    raise GateError(
                        f"fragment {path!r} unreadable at head: {exc}") from exc
                reason = validate_fragment(text)
                if reason:
                    report.append(f"HR-44 invalid fragment {path!r}: {reason}")
    except GateError as exc:
        print(f"CHANGELOG-FRAGMENT ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"CHANGELOG-FRAGMENT FAIL: {len(report)} finding(s)")
        for line in report:
            print(f"  - {line}")
        return 1
    note = f" (release consolidation of {', '.join(releases)})" if releases else ""
    print("CHANGELOG-FRAGMENT OK: every material skill edit ships a valid "
          f"fragment between {args.base!r} and {args.head!r}{note}.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
