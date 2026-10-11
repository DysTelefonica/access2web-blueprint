#!/usr/bin/env python3
# ci-pattern asset — version bump gate (HR-44, DysTelefonica/team-skills#156)
"""Version bump gate: skill edits carry a metadata.version increment (HR-44).

Takes a base ref and a head ref in the consumer's repository and gates
every skill directory touched by the diff: with no ``--skill-dir`` it
discovers each directory with a SKILL.md at head that the diff touches
(any depth, e.g. ``skills/*/`` and ``personal/*/*/``); explicit
``--skill-dir`` values pin the directories instead. For every such skill
with files changed between base and head it compares the
metadata.version fields:

- skill files changed between base and head with the same version →
  finding (HR-44: material edit without bump);
- head version < base version (semver tuple comparison) → finding
  (HR-44: downgrade);
- no skill files changed, or the skill is new (no SKILL.md at base) →
  clean.

Fail-closed: unreadable frontmatter at either ref, a base ref that does
not resolve, or a missing skill directory exits 2 — the gate never says
OK about a version it could not read (HR-3).

The comparison uses semver tuple semantics: versions are split by dots
and compared element by element, padding shorter versions with zeros
(``0.4`` ≡ ``0.4.0``).

Exit codes: 0 clean · 1 findings · 2 invalid input.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

VERSION = re.compile(r"^\s*version:\s*[\"']?(\d+(?:\.\d+)*)[\"']?\s*$", re.M)
FRONTMATTER = re.compile(r"^---\n(.+?)\n---", re.DOTALL)


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
        raise GateError(f"ref {ref!r} does not resolve (empty subject)")
    return r.stdout.strip()


def read_skill_md(repo: Path, ref: str, skill_dir: str) -> str:
    """Read SKILL.md content at the given ref."""
    return git(repo, "show", f"{ref}:{skill_dir}/SKILL.md")


def extract_version(content: str) -> str:
    m = FRONTMATTER.search(content)
    if not m:
        raise GateError("frontmatter block (--- ... ---) not found in SKILL.md")
    vm = VERSION.search(m.group(1))
    if not vm:
        raise GateError("metadata.version not found in SKILL.md frontmatter")
    return vm.group(1)


def parse_version(v: str) -> tuple:
    return tuple(int(x) for x in v.split("."))


def changed_files(repo: Path, base: str, head: str, skill_dir: str) -> list[str]:
    out = git(repo, "diff", "--name-only", f"{base}..{head}", "--", skill_dir)
    return [l for l in out.splitlines() if l.strip()]


def discover_skill_dirs(repo: Path, base: str, head: str) -> list[str]:
    """Skill dirs (a SKILL.md at head) with files touched by the diff."""
    dirs = set()
    for f in git(repo, "diff", "--name-only", f"{base}..{head}").splitlines():
        parts = f.strip().split("/")
        for i in range(1, len(parts)):
            d = "/".join(parts[:i])
            if subprocess.run(["git", "-C", str(repo), "cat-file", "-e",
                               f"{head}:{d}/SKILL.md"],
                              capture_output=True).returncode == 0:
                dirs.add(d)
    return sorted(dirs)


def main(argv):
    ap = argparse.ArgumentParser(description="Version bump gate (HR-44).")
    ap.add_argument("--repo", required=True)
    ap.add_argument("--base", required=True)
    ap.add_argument("--head", required=True)
    ap.add_argument("--skill-dir", action="append")
    args = ap.parse_args(argv)
    repo = Path(args.repo)
    report = []
    try:
        base_sha = rev_parse(repo, args.base)
        head_sha = rev_parse(repo, args.head)
        for sd in args.skill_dir or discover_skill_dirs(repo, base_sha, head_sha):
            changed = changed_files(repo, base_sha, head_sha, sd)
            if not changed:
                continue
            cat = subprocess.run(["git", "-C", str(repo), "cat-file", "-e",
                                  f"{base_sha}:{sd}/SKILL.md"], capture_output=True)
            if cat.returncode != 0:
                continue  # skill is new at head: no base version to compare
            base_version = parse_version(
                extract_version(read_skill_md(repo, base_sha, sd)))
            head_version = parse_version(
                extract_version(read_skill_md(repo, head_sha, sd)))
            if head_version < base_version:
                report.append(f"HR-44 version downgrade in {sd!r}: "
                              f"{'.'.join(map(str, base_version))} → "
                              f"{'.'.join(map(str, head_version))}")
            elif head_version == base_version:
                report.append(f"HR-44 {len(changed)} skill file(s) changed in "
                              f"{sd!r} but metadata.version stayed at "
                              f"{'.'.join(map(str, head_version))} — a material "
                              "edit requires a version bump")
    except GateError as exc:
        print(f"VERSION-BUMP ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"VERSION-BUMP FAIL: {len(report)} finding(s)")
        for line in report:
            print(f"  - {line}")
        return 1
    print("VERSION-BUMP OK: no material skill edit without a version bump "
          f"between {args.base!r} and {args.head!r}.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
