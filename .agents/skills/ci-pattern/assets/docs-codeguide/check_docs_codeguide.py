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
import glob
import importlib.util
import itertools
import re
import sys
from pathlib import Path

BACKTICK_RE = re.compile(r"`([^`]+)`")
_MARKER_RE = re.compile(r"<[^<>/]+>")
_GLOB_RE = re.compile(r"[*?\[]")

DEFAULT_CODE_ROOT = "internal/components"
DEFAULT_CODE_GLOB = "*.go"

# Prefijos que NO se tratan como rutas de repo aunque vivan entre
# comillas invertidas: cualquier URL, dirección de correo o ruta de
# archivo de sistema queda fuera.
_URL_PREFIXES = ("http://", "https://", "mailto:", "ftp://", "file://")


def _pin_output_encoding() -> None:
    """Hace la salida del gate independiente de la locale del runner."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def _load_fences():
    """Carga el módulo compartido ``assets/markdown_fences.py`` por ruta, en el mismo proceso."""
    path = Path(__file__).resolve().parent.parent / "markdown_fences.py"
    spec = importlib.util.spec_from_file_location("markdown_fences", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_FENCES = _load_fences()


def _strip_fenced_blocks(content: str) -> str:
    """Quita los bloques de código cercados antes de buscar citas.

    Un bloque aporta seis comillas invertidas; el regex de citas
    emparejaría tres y dejaría dos sueltas, invirtiendo el emparejamiento
    del código en línea posterior (falsos positivos y, peor, rutas rotas
    sin detectar). La detección (CommonMark) vive en ``markdown_fences``,
    compartida con los demás gates de documentación. Las líneas se
    conservan en blanco para no mover la numeración.
    """
    return "".join("\n" if fenced else line for _n, line, fenced in _FENCES.classify_lines(content))


def extract_path_refs(content: str) -> list[str]:
    """Devuelve las rutas de repo citadas en ``content``.

    Reglas: sólo comillas invertidas (`` `...` ``) fuera de bloques de
    código cercados (tres o más comillas invertidas o tildes); la cadena
    debe contener al menos una barra, no puede empezar por prefijo de URL
    (``http://``, ``https://``, ``mailto:``, ``ftp://``, ``file://``)
    ni por un carácter no alfabético (``/usr/bin`` queda fuera). Las
    repeticiones dentro de una página colapsan a una sola entrada.
    """
    seen: set[str] = set()
    refs: list[str] = []
    for match in BACKTICK_RE.finditer(_strip_fenced_blocks(content)):
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


def _expand_braces(token: str) -> list[str]:
    """Expande la primera alternancia ``{a,b}`` (recursivo); sin llaves válidas devuelve ``[token]``."""
    start = token.find("{")
    while start != -1:
        depth, cuts = 0, [start]
        for pos in range(start, len(token)):
            char = token[pos]
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    parts = [*cuts, pos]
                    options = [token[a + 1 : b] for a, b in itertools.pairwise(parts)]
                    if len(options) < 2:
                        break
                    head, tail = token[:start], token[pos + 1 :]
                    return [out for opt in options for out in _expand_braces(head + opt + tail)]
            elif char == "," and depth == 1:
                cuts.append(pos)
        start = token.find("{", start + 1)
    return [token]


def _glob_matches(repo: Path, pattern: str) -> bool:
    return bool(glob.glob(str(Path(glob.escape(str(repo))) / pattern), recursive=True))


def _citation_problem(repo: Path, token: str) -> str | None:
    """Motivo por el que una cita no es válida, o ``None`` si lo es (#737).

    - Ruta literal: existe.
    - Glob (``*``, ``?``, ``[``): casa con al menos un fichero o directorio.
    - Segmento marcador ``<nombre>``: se valida el prefijo anterior al
      primer marcador (que puede llevar glob); lo que sigue no se valida.
    """
    segments = token.split("/")
    marker_at = next((i for i, seg in enumerate(segments) if _MARKER_RE.search(seg)), None)
    head = "/".join(segments if marker_at is None else segments[:marker_at])
    if marker_at is not None and not head:
        return "marcador sin prefijo"
    if _GLOB_RE.search(head):
        if _glob_matches(repo, head):
            return None
        return "glob sin coincidencias" if marker_at is None else f"prefijo inexistente: {head}"
    if (repo / head).exists():
        return None
    # Cadena vacía: ruta literal inexistente (se nombra tal cual, como siempre).
    return f"prefijo inexistente: {head}" if marker_at is not None else ""


def check_citation(repo: Path, token: str) -> str | None:
    """Entrada de ``missing_in_repo`` para ``token``, o ``None`` si todas sus expansiones valen."""
    expansions = _expand_braces(token)
    problems = [(e, _citation_problem(repo, e)) for e in expansions]
    failed = [(e, p) for e, p in problems if p is not None]
    if not failed:
        return None
    if len(expansions) == 1:
        reason = failed[0][1]
        return f"{token} ({reason})" if reason else token
    details = "; ".join(f"{expanded}: {reason or 'inexistente'}" for expanded, reason in failed)
    return f"{token} ({details})"


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

    Caso de una sola raíz de :func:`check_docs_codeguide_roots`.

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
    return check_docs_codeguide_roots(repo, [(code_root, code_glob)], docs_path)


def check_docs_codeguide_roots(
    repo: Path,
    roots: list[tuple[str, str]],
    docs_path: str = "docs/codebase",
) -> dict[str, list[str]]:
    """Como :func:`check_docs_codeguide` con varias raíces (``root``, ``glob``).

    El sentido adelante (docs -> repo) se evalúa una sola vez: las citas
    pertenecen al codeguide, no a una raíz. El sentido atrás se evalúa por
    raíz; con más de una raíz cada módulo ausente se nombra como
    ``módulo (raíz)``, con una sola queda el nombre desnudo. Una lista de
    raíces vacía es un error de uso (``ValueError``): sin dato no hay
    medición.
    """
    if not roots:
        raise ValueError("at least one (code_root, code_glob) pair is required")
    if not repo.is_dir():
        raise FileNotFoundError(f"repo path is not a directory: {repo}")

    docs_text, pages = _read_docs(repo / docs_path)
    missing_in_repo: list[str] = []

    if not pages:
        # Fail-closed: el codeguide sin páginas nunca declara OK.
        missing_in_repo.append(f"<empty {docs_path}/>")
    else:
        cited: set[str] = set()
        for page in pages:
            cited.update(extract_path_refs(page.read_text(encoding="utf-8")))
        for path in sorted(cited):
            problem = check_citation(repo, path)
            if problem is not None:
                missing_in_repo.append(problem)

    missing_in_codeguide: list[str] = []
    for code_root, code_glob in roots:
        modules = _list_modules(repo / code_root, code_glob)
        if not modules:
            missing_in_codeguide.append(f"<empty {code_root}/>")
            continue
        # El nombre debe aparecer literalmente en el texto del
        # codeguide: una PR que cambie un nombre de módulo debe
        # actualizar el codeguide, mientras tanto el gate lo nombra.
        for module in modules:
            if module not in docs_text:
                missing_in_codeguide.append(
                    module if len(roots) == 1 else f"{module} ({code_root})"
                )

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
        action="append",
        default=None,
        help=(
            "Ruta relativa al repo con los módulos de primer nivel a enumerar "
            f"(default: {DEFAULT_CODE_ROOT}). Repetible: un --code-glob por raíz."
        ),
    )
    parser.add_argument(
        "--code-glob",
        action="append",
        default=None,
        help=(
            "Glob que un módulo de primer nivel debe contener para contar "
            f"(default: {DEFAULT_CODE_GLOB}). Repetible, emparejado con --code-root."
        ),
    )
    return parser


def main(argv: list[str]) -> int:
    _pin_output_encoding()
    parser = _build_parser()
    args = parser.parse_args(argv[1:])

    repo = Path(args.repo)
    roots_arg = args.code_root or [DEFAULT_CODE_ROOT]
    globs_arg = args.code_glob or [DEFAULT_CODE_GLOB] * len(roots_arg)
    if len(roots_arg) != len(globs_arg):
        print(
            f"FAIL input: --code-root ({len(roots_arg)}) and --code-glob "
            f"({len(globs_arg)}) must be given in pairs",
            file=sys.stderr,
        )
        return 2
    roots = list(zip(roots_arg, globs_arg, strict=True))
    try:
        result = check_docs_codeguide_roots(repo, roots, docs_path=args.docs_path)
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
    described = ", ".join(f"code-root={r}, glob={g}" for r, g in roots)
    print(f"docs-codeguide: OK ({pages} pages, {described})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
