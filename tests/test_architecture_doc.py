"""HR-68: docs/architecture.md must not lie about the module tree or the layer rules.

Two gates, both run as part of the normal suite:

* tree <-> code, both directions: every module described in the "Módulos" table exists under
  ``app/src/modules`` and every module on disk is described there;
* layers <-> ``scripts/check_layers.py``: the machine-readable block in the document equals
  ``ALLOWED_IMPORTS``, ``PURE_LAYERS`` and ``CROSS_CUTTING_MODULES``.

The checks take the document text, so the red fixtures below feed doctored text to the same
functions that judge the real file.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
DOC = ROOT / "docs" / "architecture.md"
MODULES_DIR = ROOT / "app" / "src" / "modules"

_spec = importlib.util.spec_from_file_location("check_layers", ROOT / "scripts" / "check_layers.py")
assert _spec and _spec.loader
check_layers = importlib.util.module_from_spec(_spec)
sys.modules["check_layers"] = check_layers
_spec.loader.exec_module(check_layers)

_MODULE_ROW = re.compile(r"^\| `([a-z_][a-z0-9_]*)` \|", re.MULTILINE)
_BLOCK = re.compile(r"```json\n(\{.*?\})\n```", re.DOTALL)


def described_modules(text: str) -> set[str]:
    """Module names in the first column of the table under '## Módulos y capas'."""
    if "## Módulos y capas" not in text:
        return set()
    section = text.split("## Módulos y capas", 1)[1].split("\n## ", 1)[0]
    return set(_MODULE_ROW.findall(section))


def modules_on_disk() -> set[str]:
    return {p.name for p in MODULES_DIR.iterdir() if p.is_dir() and not p.name.startswith("__")}


def tree_findings(text: str, on_disk: set[str]) -> list[str]:
    described = described_modules(text)
    if not described or not on_disk:
        return ["<empty>: sin módulos descritos o sin módulos en disco"]
    return [f"descrito sin existir: {m}" for m in sorted(described - on_disk)] + [
        f"existe sin describir: {m}" for m in sorted(on_disk - described)
    ]


def layer_findings(text: str) -> list[str]:
    match = _BLOCK.search(text)
    if match is None:
        return ["falta el bloque json de capas y reglas"]
    declared = json.loads(match.group(1))
    expected = {
        "layers": {k: sorted(v) for k, v in check_layers.ALLOWED_IMPORTS.items()},
        "pure_layers": sorted(check_layers.PURE_LAYERS),
        "cross_cutting_modules": sorted(check_layers.CROSS_CUTTING_MODULES),
    }
    return [
        f"{key} difiere de scripts/check_layers.py"
        for key in expected
        if declared.get(key) != expected[key]
    ]


@pytest.fixture(scope="module")
def doc_text() -> str:
    return DOC.read_text(encoding="utf-8")


def test_tree_matches_code(doc_text: str) -> None:
    assert tree_findings(doc_text, modules_on_disk()) == []


def test_layers_match_check_layers(doc_text: str) -> None:
    assert layer_findings(doc_text) == []


def test_red_module_described_but_missing(doc_text: str) -> None:
    doctored = doc_text.replace("| `expedientes` |", "| `expedientes` |\n| `fantasma` | x |", 1)
    assert "descrito sin existir: fantasma" in tree_findings(doctored, modules_on_disk())


def test_red_module_on_disk_but_undescribed(doc_text: str) -> None:
    assert "existe sin describir: nuevo" in tree_findings(doc_text, modules_on_disk() | {"nuevo"})


def test_red_empty_side_fails_closed(doc_text: str) -> None:
    assert tree_findings("", modules_on_disk())
    assert tree_findings(doc_text, set())


def test_red_rule_changed_in_doc_only(doc_text: str) -> None:
    doctored = doc_text.replace(
        '"domain": [\n      "domain"\n    ]', '"domain": [\n      "domain",\n      "ports"\n    ]'
    )
    assert doctored != doc_text
    assert "layers difiere de scripts/check_layers.py" in layer_findings(doctored)


def test_red_missing_block() -> None:
    assert layer_findings("# sin bloque") == ["falta el bloque json de capas y reglas"]
