#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/scripts/check_adoption.py
"""Fail-closed adoption preflight and scanner-image lock resolver."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tomllib
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from quality_policy import PolicyError, load_policy

COSMIC_RAY_VERSION = "8.7.0"
MUTATION_CONFIG = "cosmic-ray.toml"
DEV_LOCK = "requirements-dev.lock"
SECURITY_SOURCES = "security-images.sources.json"
SECURITY_LOCK = "security-images.lock.json"
QUALITY_POLICY = "quality-policy.json"
PYPROJECT = "pyproject.toml"
DIGEST = re.compile(r"^sha256:[0-9a-f]{64}$")

#: Pattern to detect baseline dicts/constants in Python scripts.
#: Matches `BASELINE`, `BASELINE_*`, `ALLOWLIST`, `POLICY` etc.
BASELINE_NAME_RE = re.compile(r"\b(BASELINE[A-Z_]*|ALLOWLIST[A-Z_]*|POLICY[A-Z_]*|CEILING[A-Z_]*)\b")
#: Pattern to detect shrink-only or ratchet comments in YAML/CI files.
SHRINK_ONLY_RE = re.compile(r"(shrink[_-]?only|baseline|ratchet|lock[_-]?in|shrink[-_]only)", re.IGNORECASE)
#: Patterns of ratchet gates known to be present in a repo.
RATCHET_GATES = frozenset({
    "check_ratchet.py",
    "check_ruff_ratchet.py",
    "check_mutation_sites.py",
    "check_vulture_guard.py",
})


class AdoptionError(RuntimeError):
    """A required adoption artifact is absent or unverifiable."""


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError as exc:
        raise AdoptionError(
            f"required adoption asset is unavailable: {path.name}"
        ) from exc


def _read_json_object(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(_read_text(path))
    except json.JSONDecodeError as exc:
        raise AdoptionError(f"invalid JSON in {path.name}: {exc.msg}") from exc
    if not isinstance(value, dict):
        raise AdoptionError(f"{path.name} must contain a JSON object")
    return value


def _validate_mutation_config(root: Path) -> dict[str, Any]:
    path = root / MUTATION_CONFIG
    try:
        config = tomllib.loads(_read_text(path))
    except tomllib.TOMLDecodeError as exc:
        raise AdoptionError(f"invalid TOML in {MUTATION_CONFIG}: {exc}") from exc
    section = config.get("cosmic-ray")
    if not isinstance(section, dict):
        raise AdoptionError(f"{MUTATION_CONFIG} has no [cosmic-ray] table")
    required = ("module-path", "timeout", "test-command")
    missing = [name for name in required if not section.get(name)]
    distributor = section.get("distributor")
    if not isinstance(distributor, dict) or not distributor.get("name"):
        missing.append("cosmic-ray.distributor.name")
    if missing:
        raise AdoptionError(f"{MUTATION_CONFIG} is incomplete: {', '.join(missing)}")
    return section


def _validate_quality_policy_parity(root: Path, mutation: Mapping[str, Any]) -> None:
    try:
        policy = load_policy(root / QUALITY_POLICY)
    except PolicyError as exc:
        raise AdoptionError(str(exc)) from exc
    roots = policy["subjects"]["roots"]
    if len(roots) != 1 or mutation.get("module-path") != roots[0]:
        raise AdoptionError(
            f"{MUTATION_CONFIG} module-path must match quality policy subjects.roots"
        )
    try:
        pyproject = tomllib.loads(_read_text(root / PYPROJECT))
    except tomllib.TOMLDecodeError as exc:
        raise AdoptionError(f"invalid TOML in {PYPROJECT}: {exc}") from exc
    coverage = pyproject.get("tool", {}).get("coverage", {}).get("run", {})
    if not isinstance(coverage, dict) or coverage.get("source") != roots:
        raise AdoptionError(
            f"{PYPROJECT} coverage source must match quality policy subjects.roots"
        )


def _validate_dev_lock(root: Path) -> None:
    text = _read_text(root / DEV_LOCK).replace("\\\r\n", " ").replace("\\\n", " ")
    match = re.search(
        rf"(?m)^\s*cosmic-ray=={re.escape(COSMIC_RAY_VERSION)}\b([^\n]*)", text
    )
    if match is None or not re.search(r"--hash=sha256:[0-9a-f]{64}\b", match.group(1)):
        raise AdoptionError(
            f"{DEV_LOCK} must pin cosmic-ray=={COSMIC_RAY_VERSION} with a sha256 hash"
        )


def _security_sources(root: Path) -> dict[str, str]:
    raw = _read_json_object(root / SECURITY_SOURCES)
    sources: dict[str, str] = {}
    for name, source in sorted(raw.items()):
        if (
            not isinstance(name, str)
            or not isinstance(source, str)
            or ":" not in source
        ):
            raise AdoptionError(
                f"invalid scanner source in {SECURITY_SOURCES}: {name!r}"
            )
        if (
            "replace" in source.lower()
            or "latest" in source.lower()
            or "@sha256:" in source
        ):
            raise AdoptionError(
                f"scanner source {name} must be an exact version tag: {source}"
            )
        sources[name] = source
    if set(sources) != {"gitleaks", "trivy"}:
        raise AdoptionError(
            f"{SECURITY_SOURCES} must define exactly gitleaks and trivy"
        )
    return sources


def _security_lock(root: Path, sources: Mapping[str, str]) -> dict[str, dict[str, str]]:
    raw = _read_json_object(root / SECURITY_LOCK)
    if set(raw) != set(sources):
        raise AdoptionError(f"{SECURITY_LOCK} keys do not match {SECURITY_SOURCES}")
    locked: dict[str, dict[str, str]] = {}
    for name in sorted(sources):
        entry = raw[name]
        if not isinstance(entry, dict):
            raise AdoptionError(f"invalid lock entry for {name}")
        source = entry.get("source")
        digest = entry.get("digest")
        if (
            source != sources[name]
            or not isinstance(digest, str)
            or not DIGEST.fullmatch(digest)
        ):
            raise AdoptionError(f"unresolved or mismatched scanner lock for {name}")
        locked[name] = {"source": source, "digest": digest}
    return locked


# ---- HR-15 adoption audit ----

class _BaselineFinding:
    """A discovered baseline that may or may not be covered."""
    __slots__ = ("path", "name", "source", "covered")

    def __init__(self, path: Path, name: str, source: str, covered: bool) -> None:
        self.path = path
        self.name = name
        self.source = source
        self.covered = covered

    def __repr__(self) -> str:
        state = "ok" if self.covered else "UNCOVERED"
        return f"_BaselineFinding({self.name!r}, {self.path}, {state})"


def _find_python_files(root: Path, dirs: tuple[str, ...] = ("scripts",)) -> list[Path]:
    """Yield .py files under the given sub-directory names."""
    results: list[Path] = []
    for d in dirs:
        target = root / d
        if target.is_dir():
            results.extend(p for p in target.rglob("*.py") if p.is_file())
    return results


def _detect_baseline_dicts(root: Path) -> list[_BaselineFinding]:
    """Scan Python scripts for baseline dicts/constants.

    Returns a list of findings for every file that declares a `BASELINE`-named
    dict, `ALLOWLIST`, `POLICY` or similar constant whose value is a dict literal
    (``{ ... }``) or a call that produces one.
    """
    findings: list[_BaselineFinding] = []
    for pyfile in _find_python_files(root):
        text = pyfile.read_text(encoding="utf-8")
        match = BASELINE_NAME_RE.search(text)
        if match is None:
            continue
        # Check if it's assigned as a dict constant (not just a comment)
        # Look for `NAME = { ... }` on the same or nearby line
        for line in text.splitlines():
            stripped = line.strip()
            if stripped.startswith("#"):
                continue
            if match.group() in stripped and "{" in stripped:
                findings.append(_BaselineFinding(pyfile, match.group(), "dict", False))
                break
    return findings


def _detect_shrink_only_comments(root: Path) -> list[_BaselineFinding]:
    """Scan YAML/CI files for shrink-only / ratchet declarations.

    Returns findings for every file that mentions a shrink-only or ratchet
    pattern but does not reference a known ratchet gate script.
    """
    findings: list[_BaselineFinding] = []
    for d in (".github", "ci", "workflows"):
        target = root / d
        if not target.is_dir():
            continue
        for wf in target.rglob("*.yml"):
            if not wf.is_file():
                continue
            text = wf.read_text(encoding="utf-8")
            if SHRINK_ONLY_RE.search(text):
                # Check if the file also references a ratchet gate
                has_gate = any(gate in text for gate in RATCHET_GATES)
                if not has_gate:
                    findings.append(
                        _BaselineFinding(wf, "shrink-only-comment", "yaml", False)
                    )
    return findings


def _has_check_ratchet(root: Path) -> bool:
    """Return True if the repo has at least one ratchet gate script."""
    for d in ("assets", "scripts", ".github", "ci"):
        if _has_check_ratchet_for_path(root / d):
            return True
    return False


def _has_check_ratchet_for_path(directory: Path) -> bool:
    """Return True if *directory* contains a ratchet gate script."""
    if not directory.is_dir():
        return False
    for f in directory.iterdir():
        if f.name in RATCHET_GATES and f.is_file():
            return True
    # Also check subdirectories
    for f in directory.rglob("*.py"):
        if f.name in RATCHET_GATES:
            return True
    return False


def _verify_hr15_coverage(findings: list[_BaselineFinding], root: Path) -> None:
    """Verify every discovered baseline is covered by a ratchet gate."""
    has_gate = _has_check_ratchet(root)
    uncovered: list[_BaselineFinding] = []

    for finding in findings:
        if finding.covered:
            continue
        if has_gate:
            # Heuristic 1: the script name matches a known ratchet gate
            # (e.g. check_mutation_sites.py) — it both declares AND enforces.
            script_name = finding.path.name.lower()
            if any(g.lower().replace(".py", "") in script_name for g in RATCHET_GATES):
                finding.covered = True
                continue
            # Heuristic 2: the baseline is in scripts/ and a ratchet gate
            # also lives under scripts/ — the gate operates at repo level
            # and covers all baselines in that tree.
            if finding.path.parent.is_relative_to(root / "scripts"):
                scripts_gate = _has_check_ratchet_for_path(root / "scripts")
                if scripts_gate:
                    finding.covered = True
                    continue
            # When a ratchet gate is present but the baseline is in a
            # different tree (e.g. .github/), flag as needing manual
            # verification.
            uncovered.append(finding)
        else:
            # No ratchet gate at all: every baseline is uncovered.
            uncovered.append(finding)

    if uncovered:
        msgs = []
        for f in uncovered:
            msgs.append(
                f"HR-15 baseline {f.name!r} in {f.path.relative_to(root)} "
                f"({f.source}) is NOT covered by a ratchet gate — "
                "either add check_ratchet.py (--base-baseline) or "
                "invoke an equivalent shrink-only comparison against the base branch"
            )
        raise AdoptionError(
            f"HR-15 adoption incomplete: {len(uncovered)} uncovered baseline(s). "
            + "; ".join(msgs)
        )


def _validate_hr15_adoption(root: Path) -> None:
    """Validate that all consumer baselines are covered by ratchet gates (HR-15)."""
    dict_findings = _detect_baseline_dicts(root)
    comment_findings = _detect_shrink_only_comments(root)
    all_findings = dict_findings + comment_findings
    if not all_findings:
        return  # No baselines detected — nothing to audit.
    _verify_hr15_coverage(all_findings, root)


def validate_adoption(root: Path) -> dict[str, dict[str, str]]:
    """Validate every artifact needed by mutation and security jobs."""
    mutation = _validate_mutation_config(root)
    _validate_quality_policy_parity(root, mutation)
    _validate_dev_lock(root)
    sources = _security_sources(root)
    _validate_hr15_adoption(root)
    return _security_lock(root, sources)


def resolve_security_images(
    sources: Mapping[str, str],
    *,
    runner: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run,
) -> dict[str, dict[str, str]]:
    """Resolve exact tags to registry manifest digests, failing on any ambiguity."""
    resolved: dict[str, dict[str, str]] = {}
    for name, source in sorted(sources.items()):
        command = [
            "docker",
            "buildx",
            "imagetools",
            "inspect",
            source,
            "--format",
            "{{json .Manifest.Digest}}",
        ]
        try:
            result = runner(
                command,
                capture_output=True,
                text=True,
                check=False,
                encoding="utf-8",
            )
        except OSError as exc:
            raise AdoptionError(
                "docker buildx is required to resolve scanner digests"
            ) from exc
        if result.returncode != 0:
            detail = result.stderr.strip() or f"exit {result.returncode}"
            raise AdoptionError(f"could not resolve {name} ({source}): {detail}")
        try:
            digest = json.loads(result.stdout.strip())
        except json.JSONDecodeError as exc:
            raise AdoptionError(
                f"registry returned invalid digest JSON for {name}"
            ) from exc
        if not isinstance(digest, str) or not DIGEST.fullmatch(digest):
            raise AdoptionError(
                f"registry returned an unverifiable digest for {name}: {digest!r}"
            )
        resolved[name] = {"source": source, "digest": digest}
    return resolved


def _write_lock(path: Path, lock: Mapping[str, Mapping[str, str]]) -> None:
    payload = json.dumps(lock, indent=2, sort_keys=True) + "\n"
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(payload, encoding="utf-8", newline="\n")
    temporary.replace(path)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--resolve-security-images", action="store_true")
    group.add_argument("--print-image", choices=("gitleaks", "trivy"))
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    root = args.root.resolve()
    try:
        if args.resolve_security_images:
            sources = _security_sources(root)
            lock = resolve_security_images(sources)
            _write_lock(root / SECURITY_LOCK, lock)
            print(f"wrote {SECURITY_LOCK} with {len(lock)} verified digests")
            return 0
        lock = validate_adoption(root)
        if args.print_image:
            entry = lock[args.print_image]
            print(f"{entry['source']}@{entry['digest']}")
        else:
            print("adoption preflight: complete")
        return 0
    except AdoptionError as exc:
        print(f"adoption preflight failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
