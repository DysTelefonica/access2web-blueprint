"""CI downloads images and tools only by a fixed identity (audit H5).

A tag can move; a digest cannot. An `npx` call outside the locked project resolves
the newest package from the registry on every run, so the browser build and the
E2E CLI could change between two runs of the same commit. This file keeps those
unversioned downloads out of `.github/workflows/`:

- every container image a workflow names (`services`, buildx `driver-opts image=`,
  `docker run|pull|build`) is pinned by digest;
- `npx` runs only inside `e2e/`, where `e2e/package-lock.json` pins Playwright;
- the Playwright CLI is invoked only through that lock, never a bare or `npx` copy
  resolved elsewhere.

Out of scope: `app/Dockerfile` (pinned by `test_base_image_pin.py`). A templated
reference is not literal, so it is not pinned here — but a templated *tag* is a
mutable reference and is a finding: the release image is consumed by digest
(`nombre@sha256`, #1441), never by `${image}:${GITHUB_REF_NAME}`.
"""

from __future__ import annotations

import re
from pathlib import Path

import yaml

REPO_ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = sorted((REPO_ROOT / ".github" / "workflows").glob("*.yml"))
DEPENDABOT = REPO_ROOT / ".github" / "dependabot.yml"
E2E_WORKDIR = "e2e"

DIGEST_RE = re.compile(r"@sha256:[0-9a-f]{64}")
DRIVER_IMAGE_RE = re.compile(r"\bimage=(?P<image>\S+)")
DOCKER_CMD_RE = re.compile(r"\bdocker\s+(?:run|pull|build)\b")
TAG_VALUE_RE = re.compile(r"--tag[=\s]+(?P<tag>\S+)")
# Flags whose value is a path, a name or an env pair, never an image reference.
NON_IMAGE_FLAGS = {
    "--file",
    "-f",
    "--tag",
    "-t",
    "--name",
    "--workdir",
    "-w",
    "--volume",
    "-v",
    "--publish",
    "-p",
    "--env",
    "-e",
    "--network",
    "--mount",
}
# A literal reference names a registry or namespace, or carries a tag or digest.
LITERAL_REF_RE = re.compile(
    r"^[a-z0-9][a-z0-9._/-]*(?::[A-Za-z0-9._-]+)?(?:@sha256:[0-9a-f]{64})?$"
)
PLAYWRIGHT_RE = re.compile(r"playwright\s+(?:test|install)\b")
# A templated tag: `${image}:${GITHUB_REF_NAME}` or `$IMAGE:tag`.
TEMPLATED_TAG_RE = re.compile(r":\$\{|\$[A-Z_][A-Z0-9_]*:")


def _workflows() -> list[tuple[str, dict]]:
    return [(path.name, yaml.safe_load(path.read_text(encoding="utf-8"))) for path in WORKFLOWS]


def _logical_lines(run: str) -> list[str]:
    """Join `\\`-continued lines so a `docker run` split over several lines is one command."""
    return re.sub(r"\\[ \t]*\n", " ", run).splitlines()


def _steps(doc: dict) -> list[tuple[str, dict]]:
    found = []
    for job_name, job in (doc.get("jobs") or {}).items():
        for step in job.get("steps") or []:
            found.append((job_name, step))
    return found


def _local_builds(doc: dict) -> set[str]:
    """Tags this workflow builds itself; a later `docker run` of them is not a registry pull."""
    return {
        m.group("tag")
        for _, step in _steps(doc)
        for m in TAG_VALUE_RE.finditer(str(step.get("run", "")))
    }


def _docker_image_refs(run: str, local_builds: set[str]) -> list[str]:
    """Literal image references handed to `docker run|pull|build` in one `run` script."""
    refs = []
    for line in _logical_lines(run):
        if not DOCKER_CMD_RE.search(line):
            continue
        tokens = [t.strip("\"'") for t in line.split()]
        for previous, token in zip([""] + tokens[:-1], tokens, strict=True):
            if previous in NON_IMAGE_FLAGS or token in local_builds:
                continue
            if "$" in token or not LITERAL_REF_RE.match(token):
                continue
            if "/" in token or ":" in token or "@" in token:
                refs.append(token)
    return refs


def test_every_service_image_is_pinned_by_digest() -> None:
    unpinned = []
    for name, doc in _workflows():
        for job_name, job in (doc.get("jobs") or {}).items():
            for service, spec in (job.get("services") or {}).items():
                image = str((spec or {}).get("image", ""))
                if not DIGEST_RE.search(image):
                    unpinned.append(f"{name}:{job_name}:services.{service} -> {image}")
    assert not unpinned, "service images without a digest: " + "; ".join(unpinned)


def test_every_buildx_image_option_is_pinned_by_digest() -> None:
    unpinned = []
    for name, doc in _workflows():
        for job_name, step in _steps(doc):
            for value in (step.get("with") or {}).values():
                for match in DRIVER_IMAGE_RE.finditer(str(value)):
                    if not DIGEST_RE.search(match.group("image")):
                        unpinned.append(f"{name}:{job_name} -> {match.group('image')}")
    assert not unpinned, "buildx image options without a digest: " + "; ".join(unpinned)


def test_every_docker_run_pull_or_build_image_is_pinned_by_digest() -> None:
    unpinned = []
    for name, doc in _workflows():
        builds = _local_builds(doc)
        for job_name, step in _steps(doc):
            for ref in _docker_image_refs(str(step.get("run", "")), builds):
                if not DIGEST_RE.search(ref):
                    unpinned.append(f"{name}:{job_name} -> {ref}")
    assert not unpinned, "docker commands with a mutable image tag: " + "; ".join(unpinned)


def _docker_templated_refs(run: str, local_builds: set[str]) -> list[str]:
    """Image references handed to `docker run|pull` that carry a templated tag.

    A templated tag (`name:${TAG}`) is a mutable reference: the tag can move, and
    the release image was consumed that way before the smoke had published it.
    """
    refs = []
    for line in _logical_lines(run):
        if not DOCKER_CMD_RE.search(line):
            continue
        tokens = [t.strip("\"'") for t in line.split()]
        for previous, token in zip([""] + tokens[:-1], tokens, strict=True):
            if previous in NON_IMAGE_FLAGS or token in local_builds:
                continue
            if "$" in token and TEMPLATED_TAG_RE.search(token):
                refs.append(token)
    return refs


def test_no_workflow_consumes_an_image_by_a_templated_tag() -> None:
    """An image consumed by a templated tag is consumed by a mutable reference (#1441).

    `release.yml` pulled and ran `${image}:${GITHUB_REF_NAME}`: a tag can move, and
    the E2E smoke ran before `publish` pushed it, so it could certify a stale
    artifact. The release contract is the digest (`docs/11-releases.md`).
    """
    offenders = []
    for name, doc in _workflows():
        builds = _local_builds(doc)
        for job_name, step in _steps(doc):
            for ref in _docker_templated_refs(str(step.get("run", "")), builds):
                offenders.append(f"{name}:{job_name} -> {ref}")
    assert not offenders, "images consumed by a templated tag: " + "; ".join(offenders)


def test_npx_runs_only_inside_the_locked_e2e_project() -> None:
    """`npx` outside `e2e/` downloads whatever Playwright the registry serves today."""
    offending = []
    for name, doc in _workflows():
        for job_name, step in _steps(doc):
            run = str(step.get("run", ""))
            if re.search(r"\bnpx\s", run) and step.get("working-directory") != E2E_WORKDIR:
                offending.append(f"{name}:{job_name}:{step.get('name', '<unnamed>')}")
    assert not offending, "npx outside e2e/ (unlocked download): " + "; ".join(offending)


def test_playwright_cli_comes_from_the_locked_install() -> None:
    """A bare `playwright` or an `npx` copy outside `e2e/` is not the locked version."""
    offending = []
    for name, doc in _workflows():
        for job_name, step in _steps(doc):
            for line in _logical_lines(str(step.get("run", ""))):
                for match in PLAYWRIGHT_RE.finditer(line):
                    before = line[: match.start()]
                    if before.endswith("node_modules/.bin/"):
                        continue
                    if before.endswith("npx ") and step.get("working-directory") == E2E_WORKDIR:
                        continue
                    offending.append(f"{name}:{job_name}:{step.get('name', '<unnamed>')}")
    assert not offending, "playwright outside the locked install: " + "; ".join(offending)


def test_dependabot_keeps_the_locked_playwright_current() -> None:
    """The lock that pins Playwright is only refreshed if Dependabot watches `e2e/`."""
    updates = yaml.safe_load(DEPENDABOT.read_text(encoding="utf-8"))["updates"]
    npm = [u for u in updates if u.get("package-ecosystem") == "npm"]
    assert any(u.get("directory") == "/e2e" for u in npm), "no npm update for /e2e"
