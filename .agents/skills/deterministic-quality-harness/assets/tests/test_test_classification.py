# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/tests/test_test_classification.py
"""Contract of check_test_classification: an unmarked integration file is always a violation."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "check_test_classification.py"

MODULE_MARKED = "import pytest\n\npytestmark = pytest.mark.integration\n\n\ndef test_a():\n    pass\n"
LIST_MARKED = "import pytest\n\npytestmark = [pytest.mark.slow, pytest.mark.integration]\n\n\ndef test_a():\n    pass\n"
DECORATED = (
    "import pytest\n\n\n@pytest.mark.integration\ndef test_a():\n    pass\n\n\n"
    "@pytest.mark.integration\nasync def test_b():\n    pass\n"
)
UNMARKED = "def test_a():\n    pass\n"
PARTIAL = "import pytest\n\n\n@pytest.mark.integration\ndef test_a():\n    pass\n\n\ndef test_b():\n    pass\n"
OTHER_MARK = "import pytest\n\npytestmark = pytest.mark.slow\n\n\ndef test_a():\n    pass\n"
NO_TESTS = "VALUE = 1\n"


def _run(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), "--root", str(root), *args],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def _tree(tmp_path: Path, files: dict[str, str]) -> Path:
    directory = tmp_path / "tests" / "integration"
    directory.mkdir(parents=True)
    for name, source in files.items():
        (directory / name).write_text(source, encoding="utf-8")
    return tmp_path


@pytest.mark.parametrize("source", [MODULE_MARKED, LIST_MARKED, DECORATED, NO_TESTS])
def test_marked_or_testless_file_passes(tmp_path: Path, source: str) -> None:
    result = _run(_tree(tmp_path, {"test_x.py": source}))
    assert result.returncode == 0, result.stdout
    assert "OK (1 integration file(s)" in result.stdout


@pytest.mark.parametrize("source", [UNMARKED, PARTIAL, OTHER_MARK])
def test_unmarked_file_fails_and_names_the_file(tmp_path: Path, source: str) -> None:
    result = _run(_tree(tmp_path, {"test_x.py": source}))
    assert result.returncode == 1
    assert "FAIL  test_x.py: no pytest.mark.integration marker" in result.stdout


def test_partial_file_names_only_the_unmarked_test(tmp_path: Path) -> None:
    result = _run(_tree(tmp_path, {"test_x.py": PARTIAL}))
    assert "unmarked: test_b" in result.stdout
    assert "test_a" not in result.stdout


def test_unparseable_file_fails(tmp_path: Path) -> None:
    result = _run(_tree(tmp_path, {"test_x.py": "def broken(:\n"}))
    assert result.returncode == 1
    assert "file is not parseable" in result.stdout


def test_missing_directory_fails_closed(tmp_path: Path) -> None:
    result = _run(tmp_path)
    assert result.returncode == 1
    assert "nothing was measured" in result.stdout


def test_empty_directory_fails_closed(tmp_path: Path) -> None:
    result = _run(_tree(tmp_path, {"helper.py": UNMARKED}))
    assert result.returncode == 1
    assert "no test_*.py file" in result.stdout


def test_marker_and_directory_are_configurable(tmp_path: Path) -> None:
    directory = tmp_path / "it"
    directory.mkdir()
    (directory / "test_x.py").write_text(
        "import pytest\n\npytestmark = pytest.mark.db\n\n\ndef test_a():\n    pass\n", encoding="utf-8"
    )
    assert _run(tmp_path, "--integration-dir", "it", "--marker", "db").returncode == 0
    assert _run(tmp_path, "--integration-dir", "it").returncode == 1


def test_envelope_is_well_formed_and_deterministic(tmp_path: Path) -> None:
    root = _tree(tmp_path, {"test_a.py": MODULE_MARKED, "test_b.py": UNMARKED})
    first = _run(root, "--json")
    second = _run(root, "--json")
    assert first.stdout == second.stdout
    envelope = json.loads(first.stdout)
    assert envelope["gate"] == "test_classification"
    assert envelope["status"] == "fail"
    assert envelope["indicators"] == {"unmarked_files": 1, "files_checked": 2}
    assert envelope["ceilings"] == {"unmarked_files": 0}
    assert len(envelope["findings"]) == 1
