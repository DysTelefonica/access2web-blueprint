"""The scanned base image is the base the release image is built from.

Issue #818: `security-deep.yml` scans a base image by digest and `app/Dockerfile`
builds from a base image by digest. Nothing tied the two, so the scan could keep
measuring an image the release no longer used — the same class of gap as a gate
that reports on a different artifact than the one shipped.

The digest is the contract: a tag can move, a digest cannot.
"""

from __future__ import annotations

import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
DOCKERFILE = REPO_ROOT / "app" / "Dockerfile"
SECURITY_DEEP = REPO_ROOT / ".github" / "workflows" / "security-deep.yml"

FROM_RE = re.compile(r"^FROM\s+(?P<image>[^\s@]+)@(?P<digest>sha256:[0-9a-f]{64})(?:\s+AS\s+\w+)?\s*$", re.M)
SCANNED_RE = re.compile(r"\bimage\s+(?P<image>[^\s@]+)@(?P<digest>sha256:[0-9a-f]{64})\b")


def _from_lines() -> list[tuple[str, str]]:
    return [(m.group("image"), m.group("digest")) for m in FROM_RE.finditer(DOCKERFILE.read_text())]


def _scanned() -> list[tuple[str, str]]:
    return [
        (m.group("image"), m.group("digest"))
        for m in SCANNED_RE.finditer(SECURITY_DEEP.read_text())
    ]


def test_both_dockerfile_stages_use_the_same_pinned_base() -> None:
    """Builder and runtime must not drift apart: the image is one artifact."""
    stages = _from_lines()
    assert len(stages) == 2, f"expected the two stages of app/Dockerfile, found {stages}"
    assert stages[0] == stages[1], f"builder and runtime differ: {stages}"


def test_security_deep_scans_the_base_the_release_is_built_from() -> None:
    """Every scan in the job measures the image the release ships, by digest.

    The blocking and inventory steps all name the same image, so the invariant is
    over the whole set: none of them may measure something the Dockerfile does not
    build from.
    """
    built = dict(_from_lines())
    scanned = _scanned()
    assert scanned, "security-deep.yml does not scan any pinned image"
    for image, digest in scanned:
        assert image in built, f"the scan measures {image}, which no stage of app/Dockerfile uses"
        assert built[image] == digest, (
            f"security-deep scans {image}@{digest} but app/Dockerfile builds from "
            f"{image}@{built[image]}"
        )


def test_the_scan_covers_every_severity_split_the_gate_declares() -> None:
    """A scan set without CRITICAL would pass by not looking (#818, Hard Rule 18)."""
    text = SECURITY_DEEP.read_text()
    assert "--severity CRITICAL" in text, "no step scans CRITICAL"
    assert "--severity HIGH --ignore-unfixed" in text, "no step blocks on fixable HIGH"
    assert "--severity HIGH --exit-code 0" in text, "no step keeps unfixed HIGH visible"
