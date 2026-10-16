"""Pins the governance manifest to the byte content of the files it covers.

The ``.governance-manifest.json`` declares the ``sha256`` of every artefact
the ci-pattern adoption installed in the consumer (fase 3, G3.1). The
gate ``ci-pattern adoption check --phase 3`` reads each declared file,
recomputes the hash, and compares against the manifest. A drift between
the two fails the gate naming the file: a real byte change in any of the
six ``pattern_paths`` without a re-stamp is a regression.

The test exists so the manifest is read on every CI run, not only when
the operator remembers to invoke the verificador. The hash is computed
on LF-normalised bytes, matching the convention in the ci-pattern CLI
(``assets/bin/ci-pattern::sha256_file``); CRLF on Windows is not drift.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
MANIFEST_PATH = REPO_ROOT / ".governance-manifest.json"


def _sha256_lf(path: Path) -> str:
    """sha256 of the file's content with CRLF normalised to LF (ci-pattern rule)."""
    return hashlib.sha256(path.read_bytes().replace(b"\r\n", b"\n")).hexdigest()


def test_manifest_schema_is_pinned() -> None:
    """The manifest is the contrato de identidad del patrón; its schema is 1."""
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))

    assert manifest.get("schema_version") == 1, (
        f"schema_version must be 1 (ci-pattern SKILL.md §3); got {manifest.get('schema_version')!r}"
    )
    assert "files" in manifest and isinstance(manifest["files"], dict), (
        "manifest must declare a 'files' map (ruta → sha256)"
    )


def test_manifest_files_match_the_tree() -> None:
    """For every path in manifest.files, the sha256 of the on-disk file matches.

    The same six files are the ``pattern_paths`` declared in
    ``.github/ci-pattern-adoption.json``::phases['3']. This test pins the
    manifest against the tree so a future byte change without re-stamp is
    a regression on the very next CI run, not on the next manual
    ``adoption check``.
    """
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    files: dict[str, str] = manifest["files"]

    assert files, "manifest.files is empty: nothing to pin"

    drifts: list[str] = []
    for rel, declared in files.items():
        target = REPO_ROOT / rel
        if not target.is_file():
            drifts.append(f"{rel}: not found in tree (manifest declares {declared})")
            continue
        expected = declared.removeprefix("sha256:")
        actual = _sha256_lf(target)
        if actual != expected:
            drifts.append(f"{rel}: sha256 drift — manifest {expected[:12]}… vs tree {actual[:12]}…")

    assert not drifts, (
        ".governance-manifest.json no longer matches the tree; re-stamp the manifest "
        'with `python3 -c "import hashlib, json, pathlib; ..."` or update the file '
        "if the change is intentional. Drifts:\n  - " + "\n  - ".join(drifts)
    )


def test_manifest_files_are_the_pattern_paths() -> None:
    """The manifest's ``files`` set equals the contract's ``pattern_paths`` for fase 3.

    The adoption contract declares which files the pattern installs in this
    consumer (G3.1: la sustitución sin hueco). The manifest records their
    identity. The two must be the same set; a divergence means either the
    contract drifted (una adopción silenciosa) or the manifest omitted an
    installed file.
    """
    contract = json.loads(
        (REPO_ROOT / ".github" / "ci-pattern-adoption.json").read_text(encoding="utf-8")
    )
    pattern_paths = set(contract["phases"]["3"]["pattern_paths"])
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    manifest_paths = set(manifest["files"].keys())

    missing_in_manifest = pattern_paths - manifest_paths
    extra_in_manifest = manifest_paths - pattern_paths

    assert not missing_in_manifest, (
        f"pattern_paths not covered by manifest: {sorted(missing_in_manifest)}"
    )
    assert not extra_in_manifest, (
        f"manifest files not declared as pattern_paths (residuo o entrada no inventariada): "
        f"{sorted(extra_in_manifest)}"
    )
