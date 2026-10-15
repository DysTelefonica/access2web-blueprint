#!/usr/bin/env python3
# HARNESS-PROVENANCE: access2web-blueprint issue #854 — scripts/check_pr_size.py
"""Gate del presupuesto de revisión: adiciones + borrados contra 400 líneas de autoría.

Dos modos:

    pr-event <event.json>      adiciones, borrados y etiquetas del evento pull_request;
                               con GITHUB_TOKEN en el entorno, el desglose por fichero
                               sale de la API de ficheros del PR
    local [<merge-base>]       ``git diff`` contra <merge-base> (por defecto origin/main)

Las rutas vendorizadas (``.agents/skills/**`` y ``skills/**``) son copia generada del
catálogo team-skills y no cuentan como autoría (issue #918). Si no hay desglose que
sume exactamente el total, se cuenta todo: el gate falla cerrado.

La única excepción es la etiqueta ``size:exception``, que aplica el mantenedor.
El cuerpo del PR no se lee.

Códigos de salida:
    0  dentro del presupuesto, o excepción por la etiqueta ``size:exception``
    1  fuera del presupuesto, o evento mal formado
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.request
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

MAX_CHANGED_LINES = 400
SIZE_EXCEPTION_LABEL = "size:exception"
VENDORED_ROOTS = (".agents/skills/", "skills/")
GITHUB_API = "https://api.github.com"
FILES_PER_PAGE = 100
MAX_FILE_PAGES = 30  # la API de ficheros de un PR lista como mucho 3000
REPO_PATTERN = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+")


@dataclass(frozen=True)
class FileChange:
    """Líneas añadidas y borradas de un fichero del PR. ``previous_path`` solo en renombrados."""

    path: str
    added: int
    deleted: int
    previous_path: str | None = None

    @property
    def lines(self) -> int:
        return self.added + self.deleted


def has_size_exception_label(labels: Iterable[str]) -> bool:
    """True cuando las etiquetas llevan la excepción del mantenedor."""
    return SIZE_EXCEPTION_LABEL in labels


def is_vendored(change: FileChange) -> bool:
    """Vendorizado si su ruta lo es y, en un renombrado, también la ruta anterior.

    Un renombrado desde una ruta de autoría hacia una vendorizada sigue contando.
    """
    if not change.path.startswith(VENDORED_ROOTS):
        return False
    return change.previous_path is None or change.previous_path.startswith(VENDORED_ROOTS)


def author_lines(total: int, files: list[FileChange] | None) -> int:
    """Líneas de autoría. Falla cerrado: sin desglose, o si no suma el total, cuenta todo."""
    if files is None or sum(change.lines for change in files) != total:
        return total
    return sum(change.lines for change in files if not is_vendored(change))


def compute_local_total(cwd: Path, merge_base: str) -> int:
    """Adiciones + borrados de HEAD contra ``merge_base``, como los cuenta GitHub (regla 4)."""
    shortstat = subprocess.run(
        ["git", "diff", "--shortstat", f"{merge_base}...HEAD"],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=True,
        encoding="utf-8",
    ).stdout
    return sum(int(count) for count in re.findall(r"(\d+) (?:insertion|deletion)", shortstat))


def local_files(cwd: Path, merge_base: str) -> list[FileChange] | None:
    """Desglose por fichero de ``git diff --numstat``; None si git no lo da (falla cerrado)."""
    try:
        output = subprocess.run(
            ["git", "diff", "--numstat", "--no-renames", "-z", f"{merge_base}...HEAD"],
            cwd=str(cwd),
            capture_output=True,
            check=True,
        ).stdout.decode("utf-8", errors="surrogateescape")
        return [_numstat_record(record) for record in output.split("\0") if record]
    except (OSError, ValueError, subprocess.CalledProcessError):
        return None


def _numstat_record(record: str) -> FileChange:
    # Con -z y sin renombrados cada registro es "añadidas\tborradas\truta". Binarios: "-".
    added, deleted, path = record.split("\t", 2)
    return FileChange(path, _numstat_count(added), _numstat_count(deleted))


def _numstat_count(value: str) -> int:
    return 0 if value == "-" else int(value)


def fetch_pr_files(
    repo: str,
    number: int,
    token: str,
    *,
    opener: Callable[..., Any] = urllib.request.urlopen,
) -> list[FileChange] | None:
    """Desglose por fichero desde la API de GitHub; None ante cualquier fallo (falla cerrado).

    El token solo viaja a ``api.github.com``. ``opener`` existe para los tests.
    """
    if not REPO_PATTERN.fullmatch(repo) or not token:
        return None
    files: list[FileChange] = []
    try:
        for page in range(1, MAX_FILE_PAGES + 1):
            request = urllib.request.Request(
                f"{GITHUB_API}/repos/{repo}/pulls/{number}/files"
                f"?per_page={FILES_PER_PAGE}&page={page}",
                headers={
                    "Accept": "application/vnd.github+json",
                    "Authorization": f"Bearer {token}",
                    "X-GitHub-Api-Version": "2022-11-28",
                },
            )
            with opener(request, timeout=20) as response:
                batch = json.loads(response.read().decode("utf-8"))
            if not isinstance(batch, list):
                return None
            files.extend(_file_change(item) for item in batch)
            if len(batch) < FILES_PER_PAGE:
                return files
    except (OSError, ValueError, TypeError, KeyError):
        # Falla cerrado: red, HTTP, JSON o forma inesperada cuentan el total del PR.
        return None
    return None  # más ficheros de los que lista la API: el desglose está incompleto


def _file_change(item: dict[str, Any]) -> FileChange:
    added, deleted, filename = item["additions"], item["deletions"], item["filename"]
    previous_path = item.get("previous_filename")
    if not isinstance(added, int) or not isinstance(deleted, int) or not isinstance(filename, str):
        raise TypeError("fichero del PR mal formado")
    if previous_path is not None and not isinstance(previous_path, str):
        raise TypeError("previous_filename mal formado")
    return FileChange(filename, added, deleted, previous_path)


def _pr_event(event_path: Path) -> tuple[int, list[str], str | None, int | None]:
    pull_request = json.loads(event_path.read_text(encoding="utf-8")).get("pull_request") or {}
    additions = pull_request.get("additions")
    deletions = pull_request.get("deletions")
    if not isinstance(additions, int) or not isinstance(deletions, int):
        raise ValueError(f"{event_path}: pull_request.additions/deletions must be integers")
    labels = [label["name"] for label in pull_request.get("labels") or []]
    repo = ((pull_request.get("base") or {}).get("repo") or {}).get("full_name")
    return additions + deletions, labels, repo, pull_request.get("number")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="mode")

    pr_event = subparsers.add_parser(
        "pr-event",
        help="leer adiciones, borrados y etiquetas de un evento pull_request",
    )
    pr_event.add_argument("event_json", type=Path)

    local = subparsers.add_parser("local", help="git diff --shortstat contra <merge-base>")
    local.add_argument(
        "merge_base",
        nargs="?",
        default="origin/main",
        help="referencia base (por defecto: origin/main)",
    )
    return parser


def _pin_output_encoding() -> None:
    """Fija stdout y stderr a UTF-8 (Hard Rule 8)."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = _build_parser().parse_args(argv)
    if args.mode is None:  # sin subcomando: el paso de CI actual sigue funcionando
        args.mode, args.merge_base = "local", "origin/main"

    labels: list[str] = []
    if args.mode == "local":
        total = compute_local_total(Path.cwd(), args.merge_base)
        files = local_files(Path.cwd(), args.merge_base)
    else:
        total, labels, repo, number = _pr_event(args.event_json)
        token = os.environ.get("GITHUB_TOKEN", "")
        files = None
        if token and isinstance(repo, str) and isinstance(number, int):
            files = fetch_pr_files(repo, number, token)
    counted = author_lines(total, files)

    if counted <= MAX_CHANGED_LINES or has_size_exception_label(labels):
        return 0

    print(f"Tamaño del PR: {counted} líneas (límite {MAX_CHANGED_LINES}).", file=sys.stderr)
    if counted != total:
        print(f"Excluidas {total - counted} líneas de rutas vendorizadas.", file=sys.stderr)
    elif files is None:
        print("Sin desglose por fichero: se cuenta el total del PR.", file=sys.stderr)
    if args.mode == "pr-event":
        print(
            f"Parta el PR. Solo el mantenedor puede aplicar '{SIZE_EXCEPTION_LABEL}', "
            "tras responder las cuatro preguntas del cuerpo.",
            file=sys.stderr,
        )
    else:
        print("El modo local no tiene excepción: reduzca el diff o parta el PR.", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
