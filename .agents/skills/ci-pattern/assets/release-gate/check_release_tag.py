#!/usr/bin/env python3
# ci-pattern asset — release gate: annotated semver tag on the default branch (#195)
"""Release gate over the tag that triggers a release (HR-19, HR-20).

Runs locally with ``git`` against the release repository — no network, no
host API. The subject is the tag named by ``--tag``; the consumer's release
workflow calls this gate right after checkout of that tag.

Rules:

- HR-19: the tag name MUST match ``^v(0|[1-9]\\d*)\\.(0|[1-9]\\d*)\\.(0|[1-9]\\d*)$``
  (strict semver, no leading zeros, no prerelease suffixes on the stable
  channel) and the tag object MUST be annotated, never lightweight.
- HR-19: the version MUST be greater than the last existing semver tag
  (the tagged release itself is excluded from that comparison, so the very
  first release is judgeable).
- HR-20: the tagged commit MUST be an ancestor of the default branch
  (``--default-branch``, default ``main``) — a hotfix lands on main first
  and never ships from a side branch.

Liveness: a missing tag, a repository that is not a git repository, an
empty repository or a missing default branch exits 2 — the gate refuses to
say OK about a subject it could not measure. Findings exit 1; clean exit 0.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys

SEMVER = re.compile(r"^v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def git(repo, *args):
    return subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True)


def must_git(repo, *args) -> str:
    r = git(repo, *args)
    if r.returncode != 0:
        raise GateError(f"git {' '.join(args)} failed: {r.stderr.strip()}")
    return r.stdout.strip()


def main(argv):
    ap = argparse.ArgumentParser(description="Release gate (HR-19, HR-20).")
    ap.add_argument("--repo", required=True)
    ap.add_argument("--tag", required=True)
    ap.add_argument("--default-branch", default="main")
    args = ap.parse_args(argv)
    try:
        if git(args.repo, "rev-parse", "--git-dir").returncode != 0:
            raise GateError(f"{args.repo!r} is not a git repository (empty subject)")
        if git(args.repo, "rev-parse", "-q", "--verify",
               f"refs/tags/{args.tag}").returncode != 0:
            raise GateError(f"tag {args.tag!r} does not exist in the repository "
                            "(empty subject)")
        if git(args.repo, "rev-parse", "-q", "--verify",
               f"refs/heads/{args.default_branch}").returncode != 0:
            raise GateError(f"default branch {args.default_branch!r} does not exist")

        report = []
        version = SEMVER.fullmatch(args.tag)
        if version is None:
            report.append("HR-19 tag name does not match vMAJOR.MINOR.PATCH "
                          "(strict semver: no leading zeros, no prerelease suffix)")
        tag_type = must_git(args.repo, "cat-file", "-t", f"refs/tags/{args.tag}")
        if tag_type != "tag":
            report.append("HR-19 tag is lightweight, not annotated "
                          "(create it with git tag -a -m ...)")
        commit = must_git(args.repo, "rev-parse", f"refs/tags/{args.tag}^{{}}")
        ancestor = git(args.repo, "merge-base", "--is-ancestor",
                       commit, f"refs/heads/{args.default_branch}")
        if ancestor.returncode not in (0, 1):
            raise GateError(f"git merge-base failed: {ancestor.stderr.strip()}")
        if ancestor.returncode == 1:
            report.append("HR-20 tagged commit is not reachable from the default branch "
                          f"{args.default_branch!r} — a hotfix lands on main first")
        if version is not None:
            semver = tuple(int(g) for g in version.groups())
            others = []
            for t in must_git(args.repo, "tag", "--list", "v*").split():
                if t == args.tag:
                    continue
                m = SEMVER.fullmatch(t)
                if m:
                    others.append(tuple(int(g) for g in m.groups()))
            if others and semver <= max(others):
                last = max(others)
                report.append(f"HR-19 version is not greater than the last semver tag "
                              f"v{last[0]}.{last[1]}.{last[2]} — a release always bumps")
    except GateError as exc:
        print(f"RELEASE-GATE ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"RELEASE-GATE FAIL: {len(report)} finding(s) on tag {args.tag!r}")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"RELEASE OK: tag {args.tag!r} is annotated semver, greater than the last "
          f"release, and reachable from {args.default_branch!r}.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
