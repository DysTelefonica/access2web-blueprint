#!/usr/bin/env python3
"""check_where_to_work.py — la página «dónde trabajar» (#1201).

Cada consumer declara `docs/codebase/where-to-work.md` con una fila por tarea
canónica, en la forma «lee → toca → prueba», más una fila local con el índice
del repo y sus skills obligatorias. Así un agente nuevo sabe dónde tocar sin
explorar el código.

El gate mide **presencia y deriva**, no prosa: la página existe, tiene las filas
`form`, `ci`, `schema` y `local`, cada fila declara las tres columnas y cada
ruta citada existe en el repo. Que la receta sea *correcta* lo decide el dueño
del repo; que la ruta siga existiendo, este gate.

Uso: ``python3 check_where_to_work.py <repo>``. Salida: 0 sin hallazgos, 1 con
hallazgos, 2 si la página no se puede leer (fail-closed, HR-3).
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

#: Filas que la página tiene que declarar, con su sentido.
REQUIRED_TASKS = ("form", "ci", "schema", "local")

#: Las tres columnas que cada fila nombra.
COLUMNS = ("lee", "toca", "prueba")

#: Página por defecto.
DEFAULT_PAGE = "docs/codebase/where-to-work.md"

#: Token entre acentos graves: una ruta citada o un comando.
BACKTICKED = re.compile(r"`([^`\n]+)`")

#: Un marcador de posición o un glob no es una ruta que pueda existir.
PLACEHOLDER = ("<", ">", "{", "}", "*", "?", "|")


class PageUnreadableError(RuntimeError):
    """La página existe pero no se puede leer: fallo cerrado (HR-3).

    El mensaje lo compone la clase (TRY003): quien lanza dice QUÉ no se pudo
    leer y por qué, no el texto entero.
    """

    def __init__(self, path: Path, cause: object) -> None:
        super().__init__(f"{path}: {cause}")


def _pin_output_encoding() -> None:
    """Make gate output independent from the runner locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def page_path(repo: Path) -> Path:
    """Ruta de la página dentro del repo."""
    return repo / DEFAULT_PAGE


def _read(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        raise PageUnreadableError(path, exc) from exc
    except OSError as exc:
        raise PageUnreadableError(path, exc) from exc


def _cells(line: str) -> list[str]:
    """Celdas de una fila de tabla markdown."""
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _table_rows(text: str) -> list[list[str]]:
    """Filas de tabla de la página (encabezado y cuerpo, sin separadores)."""
    rows: list[list[str]] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue
        cells = _cells(stripped)
        if all(set(cell) <= set("-: ") for cell in cells):
            continue  # separador de tabla
        rows.append(cells)
    return rows


def _key_of(cells: list[str]) -> str:
    """Clave de una fila: el token entre acentos graves de la primera celda."""
    match = BACKTICKED.search(cells[0]) if cells else None
    return match.group(1).strip() if match else ""


def _task_rows(text: str) -> dict[str, list[str]]:
    """Filas canónicas declaradas, por clave."""
    found: dict[str, list[str]] = {}
    for cells in _table_rows(text):
        key = _key_of(cells)
        if key in REQUIRED_TASKS and key not in found:
            found[key] = cells
    return found


def _has_columns(rows: list[list[str]]) -> bool:
    """Si alguna fila de tabla declara las tres columnas."""
    for cells in rows:
        names = [cell.strip().lower() for cell in cells]
        if all(column in names for column in COLUMNS):
            return True
    return False


def _cited_paths(text: str) -> list[str]:
    """Rutas citadas en la página: tokens con barra y sin espacios."""
    paths: list[str] = []
    for token in BACKTICKED.findall(text):
        candidate = token.strip()
        if not candidate or " " in candidate:
            continue
        if candidate.startswith(("http://", "https://", "mailto:")):
            continue
        if "/" not in candidate:
            continue
        if any(mark in candidate for mark in PLACEHOLDER):
            continue  # marcador de posición o glob
        if candidate not in paths:
            paths.append(candidate)
    return paths


def missing_tasks(repo: Path) -> set[str]:
    """Filas canónicas que la página no declara."""
    path = page_path(repo)
    if not path.is_file():
        return set(REQUIRED_TASKS)
    rows = _task_rows(_read(path))
    return {task for task in REQUIRED_TASKS if task not in rows}


def check(repo: Path) -> list[str]:
    """Hallazgos de la página «dónde trabajar». Vacío cuando todo cuadra."""
    path = page_path(repo)
    if not path.is_file():
        return [
            f"la página «dónde trabajar» no está: {DEFAULT_PAGE}. "
            "Un agente nuevo no tiene dónde mirar sin explorar el código."
        ]
    text = _read(path)
    findings: list[str] = []
    rows = _table_rows(text)
    if not _has_columns(rows):
        findings.append("la tabla de «dónde trabajar» no declara las columnas lee, toca y prueba")
    declared = _task_rows(text)
    for task in REQUIRED_TASKS:
        if task not in declared:
            findings.append(f"falta la fila canónica «{task}» en {DEFAULT_PAGE}")
    for task, cells in declared.items():
        for index, column in enumerate(COLUMNS, start=1):
            if index >= len(cells) or not cells[index]:
                findings.append(f"la fila «{task}» no dice qué {column}: la receta está incompleta")
    for cited in _cited_paths(text):
        if not (repo / cited).exists():
            findings.append(f"la ruta citada «{cited}» no existe: la página quedó vieja")
    return findings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Comprueba la página «dónde trabajar» de un consumer."
    )
    parser.add_argument("repo", help="raíz del repo a medir")
    args = parser.parse_args(argv)
    _pin_output_encoding()
    try:
        findings = check(Path(args.repo))
    except PageUnreadableError as exc:
        print(f"::error::no se pudo leer la página: {exc}", file=sys.stderr)
        return 2
    for finding in findings:
        print(f"::error::{finding}", file=sys.stderr)
    if findings:
        return 1
    print(f"OK: {DEFAULT_PAGE} declara las filas canónicas y sus rutas existen")
    return 0


if __name__ == "__main__":
    sys.exit(main())
