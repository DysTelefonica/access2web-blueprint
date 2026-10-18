# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/tests/test_hr15_adoption.py
"""HR-15 adoption audit tests for check_adoption.py.

Verifies that the adoption preflight detects consumer baselines declared
in scripts (``BASELINE`` dicts) and fails when they are not covered by a
ratchet gate, and passes when a gate is present.
"""

from __future__ import annotations

import json
import subprocess
import sys
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parent.parent
PREFLIGHT = ASSETS / "scripts" / "check_adoption.py"
SHA256_A = "a" * 64
SHA256_B = "b" * 64


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _write_adopted_root(root: Path) -> None:
    """Write the full set of adopted assets to *root* (standard preflight)."""
    (root / "quality-policy.json").write_bytes((ASSETS / "quality-policy.json").read_bytes())
    (root / "pyproject.toml").write_text(
        '[tool.coverage.run]\nsource = ["app"]\nomit = []\n', encoding="utf-8"
    )
    (root / "cosmic-ray.toml").write_text(
        '[cosmic-ray]\nmodule-path = "app"\ntimeout = 30.0\n'
        'excluded-modules = []\ntest-command = "python -m pytest -p no:randomly"\n\n'
        '[cosmic-ray.distributor]\nname = "local"\n',
        encoding="utf-8",
    )
    (root / "requirements-dev.lock").write_text(
        f"cosmic-ray==8.7.0 --hash=sha256:{SHA256_A}\n", encoding="utf-8"
    )
    sources = {
        "gitleaks": "zricethezav/gitleaks:v8.30.1",
        "trivy": "aquasec/trivy:0.73.0",
    }
    (root / "security-images.sources.json").write_text(
        json.dumps(sources, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    lock = {
        name: {"source": source, "digest": f"sha256:{digest}"}
        for (name, source), digest in zip(
            sorted(sources.items()), (SHA256_A, SHA256_B), strict=True
        )
    }
    (root / "security-images.lock.json").write_text(
        json.dumps(lock, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def _run(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(PREFLIGHT), "--root", str(root), *args],
        capture_output=True, text=True, check=False, encoding="utf-8",
    )


def _write_baseline_script(root: Path, name: str = "gate.py", has_ratchet_ref: bool = False) -> None:
    """Create a script in scripts/ with a BASELINE dict.

    If *has_ratchet_ref* is True the file also mentions check_ratchet.py,
    which the adoption audit treats as "covered by a ratchet gate".
    """
    scripts_dir = root / "scripts"
    scripts_dir.mkdir(exist_ok=True)
    lines = [
        '#!/usr/bin/env python3\n'
        '"""Mutation-site density gate.\n\n'
        "Uses BASELINE to tolerate over-ceiling files.\n'''",
        'from pathlib import Path',
        'from dataclasses import dataclass',
        '',
        '',
        '@dataclass(frozen=True)',
        'class BaselineEntry:',
        '    sites: int',
        '    target: int',
        '    target_date: str',
        '',
        '',
        'BASELINE: dict[str, BaselineEntry] = {}',
    ]
    if has_ratchet_ref:
        lines.append('')
        lines.append('# Uses check_ratchet.py for base-branch comparison (HR-15)')
    (scripts_dir / name).write_text('\n'.join(lines), encoding="utf-8")


def _write_ratchet_gate(root: Path) -> None:
    """Create a minimal check_ratchet.py in scripts/ (the adoption audit recognises it)."""
    scripts_dir = root / "scripts"
    scripts_dir.mkdir(exist_ok=True)
    (scripts_dir / "check_ratchet.py").write_text(
        "#!/usr/bin/env python3\n"
        "import argparse, json, sys\n"
        "ap = argparse.ArgumentParser()\n"
        "ap.parse_args()\n",
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class HRA15AdoptionTest(unittest.TestCase):
    """HR-15 adoption audit tests."""

    def test_preflight_fails_when_baseline_dict_not_covered_by_ratchet(self) -> None:
        """A repo with a BASELINE dict but no ratchet gate must fail adoption."""
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            _write_adopted_root(root)
            _write_baseline_script(root, "gate.py", has_ratchet_ref=False)
            result = _run(root)

        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("HR-15", result.stderr)
        self.assertIn("uncovered", result.stderr)

    def test_preflight_passes_when_baseline_covered_by_ratchet(self) -> None:
        """A repo with a BASELINE dict AND check_ratchet.py passes adoption."""
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            _write_adopted_root(root)
            _write_baseline_script(root, "gate.py", has_ratchet_ref=False)
            _write_ratchet_gate(root)
            result = _run(root)

        self.assertEqual(result.returncode, 0, result.stderr)

    def test_preflight_passes_without_baseline_dict(self) -> None:
        """A repo with no BASELINE dict at all still passes (nothing to audit)."""
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            _write_adopted_root(root)
            result = _run(root)

        self.assertEqual(result.returncode, 0, result.stderr)

    def test_preflight_detects_baseline_in_subdir(self) -> None:
        """A baseline in a subdirectory under scripts/ is also detected."""
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            _write_adopted_root(root)
            subdir = root / "scripts" / "mutations"
            subdir.mkdir(parents=True)
            subdir.joinpath("gate.py").write_text(
                "BASELINE_MUTATION_SITES: dict[str, int] = {}\n", encoding="utf-8"
            )
            result = _run(root)

        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("HR-15", result.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
