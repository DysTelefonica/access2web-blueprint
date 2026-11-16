#!/usr/bin/env python3
# ci-pattern asset — abre un eslabón de cadena de PRs (#1627 slice 7, regla 5)
"""Abre un eslabón de una cadena de PRs con la forma fija de la regla 5.

- Eslabón intermedio: etiqueta `chain:partial` y primera línea `Refs #N`; el cuerpo
  no puede contener palabra de cierre (`Closes`, `Fixes`, `Resolves` y variantes).
- Punta (`--tip`): primera línea `Closes #N`, sin etiqueta de cadena.
- El título nunca lleva palabra de cierre: un squash merge cierra desde el título.
- Idempotente: si ya hay un PR abierto para `--head`, lo dice y no crea otro (0).
- Como máximo 3 eslabones abiertos con `chain:partial`; el cuarto se rechaza (1).
- `--dry-run` imprime el comando `gh pr create` y no crea nada.
- Todo va por `gh`. `PRERELEASE_GH_CMD` (JSON) sustituye el ejecutable; lo usan las
  pruebas con un doble. Una salida no JSON o un gh ausente falla cerrado (2), sin traza.

Códigos de salida: 0 abierto, ya abierto o ensayo correcto; 1 rechazado por la
política; 2 entrada inválida o gh ilegible (fallo cerrado).
"""
import argparse
import json
import os
import re
import shlex
import subprocess
import sys
from pathlib import Path

CHAIN_LABEL = "chain:partial"
MAX_OPEN_LINKS = 3
FAIL_CLOSED = 2
CLOSING_RE = re.compile(r"(?i)\b(close[sd]?|fix(e[sd])?|resolve[sd]?)\b")


class GhError(Exception):
    """gh no se pudo ejecutar o no devolvió lo que el script necesita (fallo cerrado)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def gh_cmd() -> list[str]:
    raw = os.environ.get("PRERELEASE_GH_CMD")
    return json.loads(raw) if raw else ["gh"]


def gh_run(args: list[str]) -> subprocess.CompletedProcess:
    try:
        return subprocess.run(gh_cmd() + args, capture_output=True, text=True,  # noqa: S603 - fixed argv
                              encoding="utf-8", check=False)
    except OSError as exc:
        reason = f"gh no se pudo ejecutar: {exc}"
        raise GhError(reason) from exc


def gh_list(args: list[str]) -> list[dict]:
    proc = gh_run(args)
    if proc.returncode != 0:
        reason = f"gh no devuelve la lista ({' '.join(args[:2])}): {proc.stderr.strip()}"
        raise GhError(reason)
    try:
        data = json.loads(proc.stdout or "[]")
    except ValueError as exc:
        reason = "gh no devolvió JSON"
        raise GhError(reason) from exc
    if not isinstance(data, list) or not all(isinstance(item, dict) for item in data):
        reason = "gh no devolvió una lista de objetos"
        raise GhError(reason)
    return data


def parse(argv):
    p = argparse.ArgumentParser(description="Abre un eslabón de cadena de PRs (#1627 s7).")
    p.add_argument("--repo", required=True, help="owner/name del repositorio")
    p.add_argument("--issue", type=int, required=True, help="issue N que referencia el eslabón")
    p.add_argument("--head", required=True, help="rama del eslabón")
    p.add_argument("--base", default="main", help="rama base (por defecto main)")
    p.add_argument("--title", required=True)
    p.add_argument("--body-file", required=True)
    p.add_argument("--tip", action="store_true", help="es la punta: Closes #N, sin etiqueta de cadena")
    p.add_argument("--dry-run", action="store_true", help="imprime el comando; no crea nada")
    return p.parse_args(argv)


def policy_refusal(a, body_text: str) -> tuple[int, str] | None:
    """Palabra de cierre en el título, o en el cuerpo de un intermedio: se rechaza."""
    if CLOSING_RE.search(a.title):
        return 2, "fail-closed: el título no puede llevar palabra de cierre (un squash merge cierra desde el título)"
    if not a.tip and CLOSING_RE.search(body_text):
        return 2, "fail-closed: un eslabón intermedio no puede llevar palabra de cierre en el cuerpo"
    return None


def build_create(a, body_text: str) -> tuple[list[str], str]:
    first = f"Closes #{a.issue}" if a.tip else f"Refs #{a.issue}"
    create = ["pr", "create", "--repo", a.repo, "--head", a.head, "--base", a.base,
              "--title", a.title, "--body", f"{first}\n\n{body_text}"]
    if not a.tip:
        create += ["--label", CHAIN_LABEL]
    return create, first


def open_or_create(a, body_text: str) -> tuple[int, str]:
    existing = gh_list(["pr", "list", "--repo", a.repo, "--head", a.head, "--state", "open", "--json", "number,url"])
    if existing:
        where = existing[0].get("url") or f"#{existing[0].get('number')}"
        return 0, f"ya hay un PR abierto para {a.head}: {where}; no se crea otro"
    open_links = len(gh_list(["pr", "list", "--repo", a.repo, "--label", CHAIN_LABEL, "--state", "open",
                              "--json", "number"]))
    if not a.tip and open_links >= MAX_OPEN_LINKS:
        return 1, f"rechazado: ya hay {open_links} eslabones abiertos con {CHAIN_LABEL} (máximo {MAX_OPEN_LINKS})"
    create, first = build_create(a, body_text)
    if a.dry_run:
        shown = list(create)
        shown[shown.index("--body") + 1] = f"<{first} ...>"
        return 0, f"dry-run: {shlex.join(['gh', *shown])}"
    made = gh_run(create)
    if made.returncode != 0:
        return 2, f"fail-closed: gh pr create falló: {made.stderr.strip()}"
    return 0, made.stdout.strip()


def run(a) -> tuple[int, str]:
    try:
        body_text = Path(a.body_file).read_text(encoding="utf-8")
    except OSError as exc:
        return 2, f"fail-closed: no puedo leer --body-file: {exc}"
    refusal = policy_refusal(a, body_text)
    if refusal is not None:
        return refusal
    try:
        return open_or_create(a, body_text)
    except GhError as exc:
        return 2, f"fail-closed: {exc}"


def main(argv=None):
    _pin_output_encoding()
    a = parse(argv if argv is not None else sys.argv[1:])
    code, msg = run(a)
    print(msg, file=sys.stderr if code == FAIL_CLOSED else sys.stdout)
    return code


if __name__ == "__main__":
    sys.exit(main())
