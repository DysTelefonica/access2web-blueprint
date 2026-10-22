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
ADOPTION_PATH = REPO_ROOT / ".github" / "ci-pattern-adoption.json"

# Los tres gates de gobierno que BP#1040 sustituyó por el canon `ci-pattern`:
# su copia local se borró del árbol y su ejecución la toma ahora la copia
# vendorizada de `.agents/skills/ci-pattern/assets/`. Cada uno tiene que quedar
# trazado en el inventario fase 0 como `retired` —como sus hermanos en las PR
# #1041 y #1104— y ninguna superficie viva puede seguir citándolo como ruta.
GATES_REPLACED_BY_CANON = (
    "scripts/check_required_jobs.py",
    "scripts/check_issue_spec.py",
    "scripts/check_pr_size.py",
)


def _living_surfaces() -> list[Path]:
    """Markdown que un lector o un agente sigue al abrir el repo.

    Se excluyen los registros históricos (`odd/`, `openspec/`) a propósito:
    narran lo que la cadena de #1040 hizo y reescribirlos falsificaría el
    histórico. Un doc vivo nombra la ruta del canon, nunca la copia retirada.
    """
    found: list[Path] = []
    for pattern in ("*.md", "docs/**/*.md", ".github/*.md"):
        found.extend(REPO_ROOT.glob(pattern))
    return sorted(set(found))


def test_canon_replaced_gates_are_absent_from_the_tree() -> None:
    """La copia local de cada gate sustituido por el canon está realmente borrada.

    El otro lado del veredicto: una fila `retired` con el fichero todavía en el
    árbol significaría que la retirada no ocurrió y que el inventario miente.
    """
    still_there = [rel for rel in GATES_REPLACED_BY_CANON if (REPO_ROOT / rel).exists()]

    assert not still_there, (
        "BP#1040 retiró estos gates en favour de la copia vendorizada del canon, pero "
        "siguen en el árbol; o la retirada no ocurrió, o la lista debe cambiar:\n  - "
        + "\n  - ".join(still_there)
    )


def test_adoption_inventory_traces_every_retired_canon_gate() -> None:
    """Cada gate sustituido por el canon tiene su fila `retired` en la fase 0.

    Sin la fila, la retirada es invisible en el contrato de adopción: se borra
    el script y el inventario deja de hablar de él, que es como #1045 cerró la
    retirada de `scripts/check_issue_spec.py` (borró la fila `adopted` en vez de
    convertirla en `retired`, al contrario que #1041 y #1104 con sus scripts).
    """
    contract = json.loads(ADOPTION_PATH.read_text(encoding="utf-8"))
    inventory = contract["phases"]["0"]["inventory"]
    retired = {e.get("path") for e in inventory if e.get("disposition") == "retired"}

    missing = [rel for rel in GATES_REPLACED_BY_CANON if rel not in retired]

    assert not missing, (
        "el inventario fase 0 de .github/ci-pattern-adoption.json no traza como `retired` "
        "los gates que BP#1040 sustituyó por el canon; añádelos con su `path` y la nota "
        "del canon que los ejecuta:\n  - " + "\n  - ".join(missing)
    )


def test_living_docs_do_not_cite_retired_canon_gate_paths() -> None:
    """Ningún doc vivo cita la ruta de un gate que BP#1040 retiró.

    Una ruta muerta en el doc que un agente o una persona sigue es el residuo
    más caro de la migración: no rompe ningún gate, solo manda a leer un
    fichero que ya no existe (así sobrevivió `scripts/check_pr_size.py` en la
    plantilla de PR).
    """
    findings: list[str] = []
    for doc in _living_surfaces():
        text = doc.read_text(encoding="utf-8")
        for rel in GATES_REPLACED_BY_CANON:
            if rel in text:
                line_no = next(i for i, line in enumerate(text.splitlines(), 1) if rel in line)
                findings.append(f"{doc.relative_to(REPO_ROOT)}:{line_no}: cita {rel}")

    assert not findings, (
        "estos docs vivos citan rutas de gates que BP#1040 retiró; súmbranse la ruta "
        "del canon en `.agents/skills/ci-pattern/assets/`:\n  - " + "\n  - ".join(findings)
    )


def test_architecture_gate_table_does_not_point_at_a_retired_path() -> None:
    """La tabla «CI gates» de `docs/architecture.md` no lista un gate en `scripts/`.

    La tabla separa nombre y ubicación en dos columnas, así que el barrido de
    rutas completas no la ve: una fila puede seguir apuntando a `scripts/` para
    un gate que ya vive en la copia vendorizada del canon sin que ninguna otra
    comprobación lo note.
    """
    doc = REPO_ROOT / "docs" / "architecture.md"
    retired_names = {Path(rel).name for rel in GATES_REPLACED_BY_CANON}

    findings: list[str] = []
    for line in doc.read_text(encoding="utf-8").splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 2 or not cells[0].startswith("`"):
            continue
        name = cells[0].strip("` ").split(" ")[0]
        if name in retired_names and cells[1].strip("` ").strip() == "scripts/":
            findings.append(f"docs/architecture.md: {name} -> {cells[1]}")

    assert not findings, (
        "la tabla de CI gates de docs/architecture.md sigue ubicando en `scripts/` gates "
        "que BP#1040 retiró; la columna «Vive en» debe apuntar a "
        "`.agents/skills/ci-pattern/assets/`:\n  - " + "\n  - ".join(findings)
    )


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
