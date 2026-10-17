"""HR-67: docs/decisions/ keeps a valid, linked and complete record of decisions.

The checks take a directory, so the red fixtures build a doctored registry in ``tmp_path`` and
run the same function that judges the real one.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import pytest

REGISTRY = Path(__file__).resolve().parent.parent / "docs" / "decisions"
STATES = {"adoptada", "rechazada", "sustituida", "aplazada"}
_FILE = re.compile(r"^\d{4}-\d{2}-\d{2}-[a-z0-9]+(-[a-z0-9]+)*\.md$")
_FIELD = re.compile(r"^- ([a-z-]+): (.*)$")
_INDEX_LINK = re.compile(r"\]\((\d{4}-\d{2}-\d{2}-[^)]+\.md)\)")


def _fields(path: Path) -> dict[str, str]:
    return {
        m.group(1): m.group(2).strip()
        for line in path.read_text("utf-8").splitlines()
        if (m := _FIELD.match(line))
    }


def registry_findings(directory: Path) -> list[str]:
    files = sorted(p for p in directory.glob("*.md") if p.name != "README.md")
    if not files:
        return [f"{directory.name}: sin ninguna decisión"]
    out: list[str] = []
    slugs = {p.stem for p in files}
    successors: dict[str, list[str]] = {}
    for path in files:
        if not _FILE.match(path.name):
            out.append(f"{path.name}: nombre fuera de AAAA-MM-DD-<slug>.md")
        f = _fields(path)
        if f.get("estado") not in STATES:
            out.append(f"{path.name}: estado fuera de la lista")
        try:
            date.fromisoformat(f.get("fecha", ""))
        except ValueError:
            out.append(f"{path.name}: fecha no ISO")
        for key in ("decide", "motivo"):
            if not f.get(key):
                out.append(f"{path.name}: {key} vacío")
        if f.get("estado") == "aplazada" and not f.get("revision"):
            out.append(f"{path.name}: aplazada sin revision")
        if f.get("estado") == "sustituida":
            targets = [t.strip() for t in f.get("sustituida-por", "").split(",") if t.strip()]
            if not targets:
                out.append(f"{path.name}: sustituida sin sustituida-por")
            for target in targets:
                if target == path.stem:
                    out.append(f"{path.name}: sustituida por sí misma")
                elif target not in slugs:
                    out.append(f"{path.name}: sustituida-por inexistente {target}")
            successors[path.stem] = targets
    for start in successors:
        stack, seen = list(successors[start]), set()
        while stack:
            node = stack.pop()
            if node == start:
                out.append(f"{start}: ciclo de sustitución")
                break
            if node not in seen:
                seen.add(node)
                stack.extend(successors.get(node, []))
    index = directory / "README.md"
    listed = set(_INDEX_LINK.findall(index.read_text("utf-8"))) if index.exists() else set()
    out += [f"índice sin listar: {n}" for n in sorted({p.name for p in files} - listed)]
    out += [f"índice lista inexistente: {n}" for n in sorted(listed - {p.name for p in files})]
    return out


def test_real_registry_is_valid() -> None:
    assert registry_findings(REGISTRY) == []


def _write(directory: Path, name: str, **fields: str) -> None:
    body = "\n".join(f"- {k.replace('_', '-')}: {v}" for k, v in fields.items())
    (directory / name).write_text(f"# t\n\n{body}\n", "utf-8")


def _valid(directory: Path, name: str = "2026-01-01-a.md", **over: str) -> None:
    fields = {"estado": "adoptada", "fecha": "2026-01-01", "decide": "dueño", "motivo": "porque"}
    fields.update(over)
    _write(directory, name, **fields)
    (directory / "README.md").write_text("\n".join(f"]({p.name})" for p in directory.glob("2*.md")))


@pytest.mark.parametrize(
    ("over", "needle"),
    [
        ({"estado": "vigente"}, "estado fuera de la lista"),
        ({"motivo": ""}, "motivo vacío"),
        ({"decide": ""}, "decide vacío"),
        ({"fecha": "ayer"}, "fecha no ISO"),
        ({"estado": "sustituida"}, "sustituida sin sustituida-por"),
        ({"estado": "sustituida", "sustituida_por": "2026-01-01-nada"}, "inexistente"),
        ({"estado": "sustituida", "sustituida_por": "2026-01-01-a"}, "sí misma"),
        ({"estado": "aplazada"}, "aplazada sin revision"),
    ],
)
def test_red_single_decision(tmp_path: Path, over: dict[str, str], needle: str) -> None:
    _valid(tmp_path, **over)
    assert any(needle in f for f in registry_findings(tmp_path))


def test_green_substitution_chain(tmp_path: Path) -> None:
    _valid(tmp_path, "2026-01-01-b.md")
    _valid(tmp_path, "2026-01-01-a.md", estado="sustituida", sustituida_por="2026-01-01-b")
    assert registry_findings(tmp_path) == []


def test_red_substitution_cycle(tmp_path: Path) -> None:
    _valid(tmp_path, "2026-01-01-a.md", estado="sustituida", sustituida_por="2026-01-01-b")
    _valid(tmp_path, "2026-01-01-b.md", estado="sustituida", sustituida_por="2026-01-01-a")
    assert any("ciclo" in f for f in registry_findings(tmp_path))


def test_red_index_out_of_sync_both_ways(tmp_path: Path) -> None:
    _valid(tmp_path)
    (tmp_path / "README.md").write_text("](2026-01-01-fantasma.md)")
    found = registry_findings(tmp_path)
    assert "índice sin listar: 2026-01-01-a.md" in found
    assert "índice lista inexistente: 2026-01-01-fantasma.md" in found


def test_red_empty_directory_fails_closed(tmp_path: Path) -> None:
    assert registry_findings(tmp_path)


DOCS = REGISTRY.parent
_DECISION_FILE = re.compile(r"^\d{4}-\d{2}-\d{2}-d(\d+)-")
_DECISION_ROW = re.compile(r"^\| D\d+ ", re.MULTILINE)
LAST_MIGRATED = 91


def migration_findings(directory: Path, texts: dict[str, str]) -> list[str]:
    """Every D-n of the old register has its own file and no document repeats the rows."""
    present = {
        int(m.group(1)) for p in directory.glob("*.md") if (m := _DECISION_FILE.match(p.name))
    }
    out = [f"falta el fichero de D{n}" for n in range(1, LAST_MIGRATED + 1) if n not in present]
    out += [
        f"{name}: repite filas de decisión" for name, t in texts.items() if _DECISION_ROW.search(t)
    ]
    return out


def _texts() -> dict[str, str]:
    names = (
        "08-decisiones-y-preguntas-abiertas.md",
        "09-arquitectura-objetivo-y-principios.md",
        "architecture.md",
    )
    return {n: (DOCS / n).read_text("utf-8") for n in names}


def test_real_migration_is_complete_and_not_duplicated() -> None:
    assert migration_findings(REGISTRY, _texts()) == []


def test_red_missing_decision_file(tmp_path: Path) -> None:
    for n in range(1, LAST_MIGRATED + 1):
        if n != 50:
            (tmp_path / f"2026-01-01-d{n}-x.md").write_text("")
    assert migration_findings(tmp_path, {}) == ["falta el fichero de D50"]


def test_red_document_repeats_decision_rows() -> None:
    found = migration_findings(REGISTRY, {"08.md": "| D5 | texto | APROBADO |"})
    assert found == ["08.md: repite filas de decisión"]
