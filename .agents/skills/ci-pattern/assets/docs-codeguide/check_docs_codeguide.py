#!/usr/bin/env python3
# check_docs_codeguide.py — gate fail-closed para el codeguide de un repo
# (issue #528, T10, A2). Comprueba en los dos sentidos que el codeguide
# y el código están alineados. Binario (exit 0/1/2), stdlib exclusivamente,
# mismo script en local y CI.
"""Verifica que ``docs/codebase/**`` y el ``code-root`` están sincronizados.

- **Adelante (docs -> repo)**: toda ruta citada en una página del
  codeguide (``docs/codebase/**/*.md`` por defecto) existe en el repo.
- **Atrás (repo -> docs)**: todo módulo de primer nivel bajo
  ``--code-root`` que case con ``--code-glob`` aparece mencionado por
  su nombre en alguna página del codeguide.

Visibilidad: idéntico en público y en privado Free — el script es
local, no usa la API de GitHub, ni secretos, ni red. Sale 0 si ambos
sentidos pasan, 1 ante cualquier ausente (y nombra cuáles), 2 si el
directorio del repo no existe o el input es inválido.

Veredicto fail-closed: una entrada vacía en cualquier lado
(codeguide sin páginas, code-root sin módulos) es un FAIL — el
comentario de la issue lo prohíbe: «mejor que el espejo gentle-ai»,
que solo cubría ``internal/components``.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

BACKTICK_RE = re.compile(r"`([^`]+)`")

# Prefijos que NO se tratan como rutas de repo aunque vivan entre
# comillas invertidas: cualquier URL, dirección de correo o ruta de
# archivo de sistema queda fuera.
_URL_PREFIXES = ("http://", "https://", "mailto:", "ftp://", "file://")


def _pin_output_encoding() -> None:
    """Hace la salida del gate independiente de la locale del runner."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


def extract_path_refs(content: str) -> list[str]:
    """Devuelve las rutas de repo citadas en ``content``.

    Reglas: sólo comillas invertidas (`` `...` ``); la cadena debe
    contener al menos una barra, no puede empezar por prefijo de URL
    (``http://``, ``https://``, ``mailto:``, ``ftp://``, ``file://``)
    ni por un carácter no alfabético (``/usr/bin`` queda fuera). Las
    repeticiones dentro de una página colapsan a una sola entrada.
    """
    seen: set[str] = set()
    refs: list[str] = []
    for match in BACKTICK_RE.finditer(content):
        token = match.group(1).strip()
        if "/" not in token:
            continue
        if token.startswith(_URL_PREFIXES):
            continue
        if not token[0].isalpha():
            continue
        if token not in seen:
            seen.add(token)
            refs.append(token)
    return refs


def _read_docs(docs_abs: Path) -> tuple[str, list[Path]]:
    """Lee todas las páginas Markdown bajo ``docs_abs``.

    Devuelve el texto concatenado (vacío si no hay páginas) y la lista
    de ficheros. Si el directorio no existe o está vacío, devuelve
    ``("", [])`` — el llamador decide si eso es un FAIL.
    """
    if not docs_abs.is_dir():
        return "", []
    pages = sorted(p for p in docs_abs.rglob("*.md") if p.is_file())
    if not pages:
        return "", []
    parts = []
    for page in pages:
        parts.append(page.read_text(encoding="utf-8"))
    return "\n".join(parts), pages


def _list_modules(code_abs: Path, code_glob: str) -> list[str]:
    """Lista los módulos de primer nivel bajo ``code_abs``.

    Un módulo cuenta si:

    - Es un fichero cuyo nombre case con ``code_glob`` (módulos
      "planos" estilo ``src/modules/users.py``); el nombre que se
      reporta es el ``stem``.
    - O es un directorio que contiene al menos un fichero que case
      con ``code_glob`` (módulos "carpeta" estilo
      ``internal/components/<name>/*.go``); el nombre que se reporta
      es el nombre del directorio.

    Devuelve los nombres en orden estable (lexicográfico) para que el
    veredicto sea reproducible.
    """
    if not code_abs.is_dir():
        return []
    modules: list[str] = []
    for entry in sorted(code_abs.iterdir()):
        if entry.is_file():
            if entry.match(code_glob):
                modules.append(entry.stem)
            continue
        if entry.is_dir() and any(entry.glob(code_glob)):
            modules.append(entry.name)
    return modules


def check_docs_codeguide(
    repo: Path,
    docs_path: str = "docs/codebase",
    code_root: str = "internal/components",
    code_glob: str = "*.go",
) -> dict[str, list[str]]:
    """Falla cerrada: devuelve el dict de ausentes; dos listas vacías = OK.

    Estructura del veredicto::

        {
            "missing_in_repo": ["internal/ghost/foo.go", ...],
            "missing_in_codeguide": ["orphan_module", ...],
        }

    Las listas vacías en ambas claves significan «OK»; cualquier
    entrada ausente en cualquier lado es un FAIL. La entrada
    ``"<empty docs/codebase/>"`` o ``"<empty internal/components/>"``
    nombra el lado vacío cuando uno de los dos está sin contenido.
    """
    if not repo.is_dir():
        raise FileNotFoundError(f"repo path is not a directory: {repo}")

    docs_abs = repo / docs_path
    code_abs = repo / code_root

    docs_text, pages = _read_docs(docs_abs)
    missing_in_repo: list[str] = []

    if not pages:
        # Fail-closed: el codeguide sin páginas nunca declara OK.
        missing_in_repo.append(f"<empty {docs_path}/>")
    else:
        cited: set[str] = set()
        for page in pages:
            cited.update(extract_path_refs(page.read_text(encoding="utf-8")))
        for path in sorted(cited):
            if not (repo / path).exists():
                missing_in_repo.append(path)

    modules = _list_modules(code_abs, code_glob)
    missing_in_codeguide: list[str] = []
    if not modules:
        missing_in_codeguide.append(f"<empty {code_root}/>")
    else:
        # El nombre debe aparecer literalmente en el texto del
        # codeguide: una PR que cambie un nombre de módulo debe
        # actualizar el codeguide, mientras tanto el gate lo nombra.
        for module in modules:
            if module not in docs_text:
                missing_in_codeguide.append(module)

    return {
        "missing_in_repo": missing_in_repo,
        "missing_in_codeguide": missing_in_codeguide,
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="check_docs_codeguide.py",
        description=(
            "Comprueba en los dos sentidos que el codeguide y el "
            "código de un repo están sincronizados (#528, T10)."
        ),
    )
    parser.add_argument(
        "repo",
        help="Ruta al directorio raíz del repo a validar.",
    )
    parser.add_argument(
        "--docs-path",
        default="docs/codebase",
        help="Ruta relativa al repo donde vive el codeguide (default: docs/codebase).",
    )
    parser.add_argument(
        "--code-root",
        default="internal/components",
        help="Ruta relativa al repo con los módulos de primer nivel a enumerar.",
    )
    parser.add_argument(
        "--code-glob",
        default="*.go",
        help="Glob que un módulo de primer nivel debe contener para contar.",
    )
    return parser


def main(argv: list[str]) -> int:
    _pin_output_encoding()
    parser = _build_parser()
    args = parser.parse_args(argv[1:])

    repo = Path(args.repo)
    try:
        result = check_docs_codeguide(
            repo,
            docs_path=args.docs_path,
            code_root=args.code_root,
            code_glob=args.code_glob,
        )
    except FileNotFoundError as exc:
        print(f"FAIL input: {exc}", file=sys.stderr)
        return 2

    missing_in_repo = result["missing_in_repo"]
    missing_in_codeguide = result["missing_in_codeguide"]

    if missing_in_repo or missing_in_codeguide:
        for path in missing_in_repo:
            print(f"FAIL docs -> repo: missing {path}", file=sys.stderr)
        for module in missing_in_codeguide:
            print(f"FAIL repo -> docs: missing {module}", file=sys.stderr)
        return 1

    pages = sum(1 for _ in (repo / args.docs_path).rglob("*.md"))
    print(
        f"docs-codeguide: OK "
        f"({pages} pages, code-root={args.code_root}, glob={args.code_glob})"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))