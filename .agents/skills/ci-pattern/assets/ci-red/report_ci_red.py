#!/usr/bin/env python3
"""ci-red: el run programado en rojo sobre la rama por defecto tiene dueño (HR-89).

Un nightly es la red de seguridad de la rama por defecto; si su rojo no abre
nada, nadie lo mira (#1508: en ardelperal/APAP_WEB, seis noches seguidas en
rojo sin una sola issue). Este asset se copia en el consumer y lo invoca un
paso del workflow programado; nunca invoca la CLI del patrón (lección #342).

Contrato:
  - Solo actúa en un run ``schedule`` sobre la rama por defecto. Con otro
    evento o en otra rama dice «nada que hacer», sale 0 y no llama a ``gh``.
  - ``--outcome failure``: si hay una issue abierta con la etiqueta ``ci-red``
    y el título ``[ci-red] <workflow>``, comenta en ella la evidencia del run;
    si no, la abre con esa etiqueta, asignada al ``coordinator`` del contrato
    del host. Una issue por workflow, no una por noche.
  - ``--outcome success``: si esa issue está abierta, comenta el run en verde
    y la cierra; si no, no hace nada.
  - El contrato del host ``MUST`` declarar ``coordinator`` (login no vacío) y
    la etiqueta ``ci-red`` (HR-34); si no, sale 2 sin llamar a ``gh``: un
    aviso sin dueño es el hueco que esta regla cierra.

Códigos de salida: 0 aviso dejado o nada que hacer · 1 ``gh`` no pudo dejar
el aviso (fail-loud, HR-3) · 2 uso, identidad del run o contrato inservibles.

Variables de entorno (las pone GitHub Actions): GITHUB_REPOSITORY,
GITHUB_EVENT_NAME, GITHUB_REF_NAME, GITHUB_WORKFLOW, GITHUB_SERVER_URL,
GITHUB_RUN_ID y GITHUB_SHA. El token de ``gh`` (``GH_TOKEN``) necesita
``issues: write``.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

LABEL = "ci-red"
EXIT_OK, EXIT_GH, EXIT_USAGE = 0, 1, 2


class UsageError(Exception):
    """Entrada inservible: sale 2 sin tocar el host."""


class GhError(Exception):
    """``gh`` no pudo dejar el aviso: sale 1."""


def load_contract(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        msg = f"host-contract ilegible ({path}): {exc}"
        raise UsageError(msg) from exc
    if not isinstance(data, dict):
        msg = f"host-contract no es un objeto JSON ({path})"
        raise UsageError(msg)
    return data


def coordinator_of(contract: dict) -> str:
    """Login del coordinador declarado; exige también la etiqueta ``ci-red``."""
    coordinator = contract.get("coordinator")
    if not isinstance(coordinator, str) or not coordinator.strip():
        msg = ("host-contract sin 'coordinator' (login no vacío): la issue ci-red "
               "no tendría dueño (HR-89)")
        raise UsageError(msg)
    labels = contract.get("labels")
    names = {entry.get("name") for entry in labels if isinstance(entry, dict)} \
        if isinstance(labels, list) else set()
    if LABEL not in names:
        msg = (f"host-contract no declara la etiqueta '{LABEL}' en 'labels' "
               "(etiqueta de gobierno, HR-34)")
        raise UsageError(msg)
    return coordinator.strip()


def gh(*args: str) -> str:
    exe = shutil.which("gh")
    if exe is None:
        msg = "gh no está en el PATH; el aviso ci-red no se puede dejar"
        raise GhError(msg)
    proc = subprocess.run([exe, *args], capture_output=True, text=True,  # noqa: S603 — lista de args, shell desactivado.
                          encoding="utf-8", check=False)
    if proc.returncode != 0:
        msg = f"gh {' '.join(args[:2])} falló (rc={proc.returncode}): {proc.stderr.strip()}"
        raise GhError(msg)
    return proc.stdout


def open_issue_number(repo: str, title: str) -> int | None:
    out = gh("issue", "list", "--repo", repo, "--state", "open", "--label", LABEL,
             "--limit", "200", "--json", "number,title")
    try:
        rows = json.loads(out or "[]")
    except ValueError as exc:
        msg = f"gh issue list devolvió algo que no es JSON: {exc}"
        raise GhError(msg) from exc
    numbers = [row["number"] for row in rows
               if isinstance(row, dict) and row.get("title") == title]
    return min(numbers) if numbers else None


def with_body(text: str, *args: str) -> str:
    """Ejecuta ``gh`` con el cuerpo en un fichero temporal (``--body-file``)."""
    handle, path = tempfile.mkstemp(prefix="ci-red-", suffix=".md")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as fh:
            fh.write(text)
        return gh(*args, "--body-file", path)
    finally:
        Path(path).unlink()


def run_identity(args: argparse.Namespace) -> dict:
    ident = {
        "repo": args.repo or os.environ.get("GITHUB_REPOSITORY", ""),
        "workflow": args.workflow or os.environ.get("GITHUB_WORKFLOW", ""),
        "run_id": os.environ.get("GITHUB_RUN_ID", ""),
        "sha": os.environ.get("GITHUB_SHA", ""),
        "server": os.environ.get("GITHUB_SERVER_URL", "https://github.com"),
    }
    for key, var in (("repo", "GITHUB_REPOSITORY"), ("workflow", "GITHUB_WORKFLOW")):
        if not ident[key].strip():
            msg = f"falta {var}: no se sabe qué run avisar"
            raise UsageError(msg)
    run_url = args.run_url
    if not run_url:
        run_url = (f"{ident['server']}/{ident['repo']}/actions/runs/{ident['run_id']}"
                   if ident["run_id"] else "sin URL del run")
    ident["run_url"] = run_url
    return ident


def evidence(ident: dict, branch: str) -> str:
    return "\n".join([
        f"- Workflow: `{ident['workflow']}`",
        f"- Rama: `{branch}`",
        f"- Commit: `{ident['sha'] or 'sin SHA'}`",
        f"- Run: {ident['run_url']}",
    ])


def report_failure(ident: dict, branch: str, coordinator: str) -> str:
    title = f"[{LABEL}] {ident['workflow']}"
    number = open_issue_number(ident["repo"], title)
    if number is not None:
        with_body(f"El run programado volvió a fallar.\n\n{evidence(ident, branch)}\n",
                  "issue", "comment", str(number), "--repo", ident["repo"])
        return f"ci-red: comentario en la issue #{number} ({title})"
    body = (
        f"El run programado de `{ident['workflow']}` falló sobre la rama por defecto. "
        "Esta issue es el aviso (HR-89): sigue abierta, y cada noche en rojo la "
        "comenta, hasta que un run programado en verde la cierre.\n\n"
        f"{evidence(ident, branch)}\n")
    url = with_body(body, "issue", "create", "--repo", ident["repo"], "--title", title,
                    "--label", LABEL, "--assignee", coordinator).strip()
    number_text = url.rsplit("/", 1)[-1] if url else "?"
    return (f"ci-red: issue #{number_text} abierta para el workflow «{ident['workflow']}», "
            f"asignada a {coordinator}")


def report_success(ident: dict, branch: str) -> str:
    title = f"[{LABEL}] {ident['workflow']}"
    number = open_issue_number(ident["repo"], title)
    if number is None:
        return f"ci-red: sin issue abierta para «{ident['workflow']}»; nada que cerrar"
    with_body(f"El run programado volvió a verde; se cierra.\n\n{evidence(ident, branch)}\n",
              "issue", "comment", str(number), "--repo", ident["repo"])
    gh("issue", "close", str(number), "--repo", ident["repo"])
    return f"ci-red: issue #{number} cerrada ({title} en verde)"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Abre, comenta o cierra la issue ci-red de un run programado (HR-89).")
    parser.add_argument("--outcome", required=True, choices=("failure", "success"))
    parser.add_argument("--default-branch", required=True,
                        help="rama por defecto (${{ github.event.repository.default_branch }})")
    parser.add_argument("--host-contract", default=".github/host-contract.json", type=Path)
    parser.add_argument("--repo", help="por defecto GITHUB_REPOSITORY")
    parser.add_argument("--workflow", help="por defecto GITHUB_WORKFLOW")
    parser.add_argument("--run-url", help="por defecto, derivada de GITHUB_SERVER_URL y GITHUB_RUN_ID")
    return parser.parse_args(argv)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the same output on a Windows desk and a Linux runner."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    event = os.environ.get("GITHUB_EVENT_NAME", "")
    branch = os.environ.get("GITHUB_REF_NAME", "")
    if event != "schedule" or branch != args.default_branch:
        print(f"ci-red: nada que hacer (evento «{event}», rama «{branch}»; HR-89 solo "
              f"actúa en un run schedule sobre «{args.default_branch}»)")
        return EXIT_OK
    try:
        ident = run_identity(args)
        coordinator = coordinator_of(load_contract(args.host_contract))
        if args.outcome == "failure":
            print(report_failure(ident, branch, coordinator))
        else:
            print(report_success(ident, branch))
    except UsageError as exc:
        print(f"ci-red: ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except GhError as exc:
        print(f"ci-red: ERROR: {exc}", file=sys.stderr)
        return EXIT_GH
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
