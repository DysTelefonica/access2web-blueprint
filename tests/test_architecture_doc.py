"""HR-68: docs/architecture.md must not lie about the module tree or the layer rules.

Two checks, both run as part of the normal suite:

* tree <-> code, both directions: every layer folder described in the "Árbol de módulos" table
  exists under ``app/src/modules/<module>/`` and every layer folder on disk is described there;
* layers <-> ``scripts/check_layers.py``: the ``architecture`` block in the document equals
  ``ALLOWED_IMPORTS`` (without the self edge), ``PURE_LAYERS`` with ``FORBIDDEN_IN_PURE_LAYERS`` and
  ``CROSS_CUTTING_MODULES``, and its ``independent`` group lists every application module on disk.

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

_TREE_ROW = re.compile(r"^\| `([^`]+)` \|", re.MULTILINE)
_BLOCK = re.compile(r"```architecture\n(\{.*?\})\n```", re.DOTALL)


def described_modules(text: str) -> set[str]:
    """Module paths in the first column of the table under '## Árbol de módulos'."""
    if "## Árbol de módulos" not in text:
        return set()
    section = text.split("## Árbol de módulos", 1)[1].split("\n## ", 1)[0]
    return set(_TREE_ROW.findall(section))


def modules_on_disk() -> set[str]:
    """Layer folders (with a Python file directly inside) of every application module."""
    found = set()
    for app in MODULES_DIR.iterdir():
        if not app.is_dir() or app.name.startswith("__"):
            continue
        for layer in app.iterdir():
            if layer.is_dir() and not layer.name.startswith("__") and any(layer.glob("*.py")):
                found.add(layer.relative_to(ROOT).as_posix())
    return found


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
        return ["falta el bloque architecture de capas y reglas"]
    declared = json.loads(match.group(1))
    expected = {
        "allowed": {
            k: sorted(set(v) - {k}) for k, v in check_layers.ALLOWED_IMPORTS.items()
        },
        "pure": {
            "layers": sorted(check_layers.PURE_LAYERS),
            "forbidden_externals": sorted(check_layers.FORBIDDEN_IN_PURE_LAYERS),
        },
        "cross_cutting_modules": sorted(check_layers.CROSS_CUTTING_MODULES),
    }
    findings = [
        f"{key} difiere de scripts/check_layers.py"
        for key in expected
        if declared.get(key) != expected[key]
    ]
    apps = sorted(
        p.relative_to(ROOT).as_posix()
        for p in MODULES_DIR.iterdir()
        if p.is_dir() and not p.name.startswith("__")
    )
    groups = [sorted(g.get("modules", [])) for g in declared.get("independent", [])]
    if groups != [apps]:
        findings.append("independent difiere de las aplicaciones de app/src/modules")
    return findings


@pytest.fixture(scope="module")
def doc_text() -> str:
    return DOC.read_text(encoding="utf-8")


def test_tree_matches_code(doc_text: str) -> None:
    assert tree_findings(doc_text, modules_on_disk()) == []


def test_layers_match_check_layers(doc_text: str) -> None:
    assert layer_findings(doc_text) == []


def test_red_module_described_but_missing(doc_text: str) -> None:
    ghost = "| `app/src/modules/fantasma/domain` | x |"
    doctored = doc_text.replace("## Árbol de módulos\n", f"## Árbol de módulos\n\n{ghost}", 1)
    assert "descrito sin existir: app/src/modules/fantasma/domain" in tree_findings(
        doctored, modules_on_disk()
    )


def test_red_module_on_disk_but_undescribed(doc_text: str) -> None:
    assert "existe sin describir: app/src/modules/nuevo/domain" in tree_findings(
        doc_text, modules_on_disk() | {"app/src/modules/nuevo/domain"}
    )


def test_red_empty_side_fails_closed(doc_text: str) -> None:
    assert tree_findings("", modules_on_disk())
    assert tree_findings(doc_text, set())


def test_red_rule_changed_in_doc_only(doc_text: str) -> None:
    doctored = doc_text.replace(
        '"domain": []', '"domain": [\n      "ports"\n    ]', 1
    )
    assert doctored != doc_text
    assert "allowed difiere de scripts/check_layers.py" in layer_findings(doctored)


def test_red_missing_block() -> None:
    assert layer_findings("# sin bloque") == ["falta el bloque architecture de capas y reglas"]


def test_red_independent_group_missing_an_application(doc_text: str) -> None:
    doctored = doc_text.replace('"app/src/modules/expedientes"', '"app/src/modules/otra"', 1)
    assert doctored != doc_text
    assert "independent difiere de las aplicaciones de app/src/modules" in layer_findings(doctored)
