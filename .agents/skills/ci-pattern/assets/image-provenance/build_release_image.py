#!/usr/bin/env python3
# ci-pattern asset — build a release image from a fresh clone of its tag (#1389)
"""Build the image of a release tag, and only from that tag.

The registry is unreachable from CI, so a person builds the image. This
script makes the release build the only easy path: it clones the remote
fresh into a temporary directory (never a working copy, never an export),
checks out the tag and refuses when

- the tag does not exist on the remote;
- the tag commit is not on the default branch;
- the checkout is dirty, or its deployable tree differs from the git
  content of the tag (a ``.gitattributes`` eol rule or a filter would put
  in the image bytes that the release manifest does not hold).

Then it runs ``<engine> build`` with the OCI labels the deploy gate reads:
``org.opencontainers.image.revision`` = the tag commit and
``org.opencontainers.image.version`` = the tag. ``--push`` also pushes.

Git runs with ``core.autocrlf=false`` per invocation (never written to any
configuration), so the checkout carries the git bytes.

Exit codes: 0 built · 1 refused (findings) · 2 unmeasurable or tool failure.
stdlib only.
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import image_provenance as ip  # noqa: E402

NO_AUTOCRLF = ("-c", "core.autocrlf=false")


def disk_entries(root: Path, exclude: list[str]) -> dict[str, str]:
    """The checked-out deployable tree, in the manifest's notation."""
    entries = {}
    for dirpath, dirnames, filenames in os.walk(root):
        here = Path(dirpath)
        for name in filenames + [d for d in dirnames if (here / d).is_symlink()]:
            path = here / name
            rel = path.relative_to(root).as_posix()
            if ip.is_excluded(rel, exclude):
                continue
            entries[rel] = ("symlink:" + str(path.readlink())) if path.is_symlink() \
                else ip.sha256(path.read_bytes())
    return entries


def clone_tag(remote: str, tag: str, cfg: dict, clone: Path) -> tuple[str | None, list[str]]:
    """Fresh clone at the tag; returns the tag commit and the refusals."""
    branch, tree, exclude = cfg["default_branch"], cfg["tree"], cfg["exclude"]
    ip.git(clone.parent, *NO_AUTOCRLF, "clone", "--quiet", "--no-checkout", "--",
           remote, str(clone))
    rc, commit = ip.git_status(clone, "rev-parse", "-q", "--verify", f"refs/tags/{tag}^{{commit}}")
    if rc != 0:
        return None, [f"tag {tag!r} does not exist on the remote: push the release tag first"]
    rc, _ = ip.git_status(clone, "merge-base", "--is-ancestor", commit,
                          f"refs/remotes/origin/{branch}")
    if rc not in (0, 1):
        msg = f"cannot read the default branch {branch!r} of the remote"
        raise ip.GateError(msg)
    refusals = []
    if rc == 1:
        refusals.append(f"tag {tag!r} ({commit[:12]}) is not on {branch!r}: "
                        "a release ships from the default branch only")
    ip.git(clone, *NO_AUTOCRLF, "checkout", "--quiet", "--detach", commit)
    dirty = ip.git(clone, *NO_AUTOCRLF, "status", "--porcelain").decode().strip()
    if dirty:
        refusals.append(f"the fresh checkout is dirty: {dirty.splitlines()[0]}")
    expected = ip.git_tree_entries(clone, commit, tree, exclude)
    actual = disk_entries(clone / tree, exclude)
    refusals += [f"checkout differs from the git content of the tag ({f})"
                 for f in ip.compare(expected, actual, {})]
    return commit, refusals


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv) -> int:
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Build a release image from its tag (#1389).")
    ap.add_argument("--config", required=True)
    ap.add_argument("--tag", required=True)
    ap.add_argument("--git-remote", help="remote URL (default: the GitHub repository)")
    ap.add_argument("--push", action="store_true", help="push the image after building")
    args = ap.parse_args(argv)
    try:
        cfg = ip.load_config(args.config)
        remote = args.git_remote or f"https://github.com/{cfg['github_repo']}.git"
        image = f"{cfg['image']}:{args.tag}"
        engine = cfg.get("engine_cmd", ["podman"])
        with tempfile.TemporaryDirectory(prefix="release-build-") as tmp:
            clone = Path(tmp) / "src"
            commit, refusals = clone_tag(remote, args.tag, cfg, clone)
            if refusals:
                print(f"RELEASE-BUILD REFUSED: {len(refusals)} finding(s) for tag {args.tag!r}")
                for line in refusals:
                    print(f"  - {line}")
                return 1
            ip.run_cmd([*engine, "build", "-f", str(clone / cfg.get("containerfile", "Dockerfile")),
                        "--label", f"{ip.ImageTree.REVISION}={commit}",
                        "--label", f"{ip.ImageTree.VERSION}={args.tag}",
                        "--label", "org.opencontainers.image.source="
                                   f"https://github.com/{cfg['github_repo']}",
                        "-t", image, str(clone)], f"building {image}")
        if args.push:
            ip.run_cmd([*engine, "push", image], f"pushing {image}")
    except ip.GateError as exc:
        print(f"RELEASE-BUILD ERROR: {exc}", file=sys.stderr)
        return 2
    print(f"RELEASE-BUILD OK: {image} built from tag {args.tag!r} ({commit[:12]})"
          f"{' and pushed' if args.push else ''}. Before `oc set image`, run: "
          f"image_provenance.py gate --config {args.config} --image {image} "
          "--current-from-cluster")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
