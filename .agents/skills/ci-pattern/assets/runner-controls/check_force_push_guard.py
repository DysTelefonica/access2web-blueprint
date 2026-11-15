#!/usr/bin/env python3
# ci-pattern asset — force-push / deletion guard (DysTelefonica/team-skills#303)
"""Detector de force-push y de borrado de ramas protegidas por contrato.

Lee el evento `push` (`--event`) y el contrato (`--contract`): si la rama
de `ref` está en `protected_branches` y el evento trae `forced: true` o
`deleted: true`, abre una issue de incidente con la evidencia (ref,
`before` -> `after`, pusher, mensaje del commit de cabeza) y devuelve 1.
Un push ordinario o una rama no protegida devuelve 0 sin llamar a la API;
cualquier evento malformado o fallo de `gh` falla cerrado con 2 (HR-3).
Runner-enforced (HR-34); el resto de la rama protegida lo audita el
detective de HR-49. Solo stdlib; cada llamada a `gh api` lleva timeout.

#1266: el merge gobernado despacha este guard por `workflow_dispatch` (#817),
y ese evento no trae `forced`/`deleted` ni el rango de nivel superior: el rango
viaja en `inputs`. El rango se lee con el lector compartido `push_range`
(#933) —el mismo que usa el detective post-push—, de modo que hay una sola
implementación de «el rango de este evento». Un dispatch no trae las banderas
y su ausencia se dice en la salida: no es una medición que haya pasado. Un
dispatch con el rango mal formado falla cerrado con 2, como un push con el
rango mal formado.
"""
import argparse
import sys

from check_push_compliance import (
    FailClosedError,
    fail,
    gh_open_incident,
    load_json,
    protected_branches,
    push_range,
)


def event_identity(event: dict) -> tuple[str, dict]:
    """Rama y repositorio del evento; cualquier forma dudosa es fallo cerrado (HR-3)."""
    ref = event.get("ref")
    if not isinstance(ref, str) or not ref.startswith("refs/heads/"):
        fail("event has no 'refs/heads/' ref to audit")
    repo = event.get("repository")
    if not isinstance(repo, dict) or not isinstance(repo.get("full_name"), str) \
            or repo["full_name"].count("/") != 1:
        fail("event has no OWNER/REPO 'repository.full_name'")
    return ref[len("refs/heads/"):], repo


def is_dispatch(event: dict) -> bool:
    """El dispatch del merge gobernado (#817) lleva el rango auditado en `inputs`."""
    return "inputs" in event


def event_flags(event: dict) -> tuple[bool, bool] | None:
    """`forced`/`deleted` del evento; `None` cuando el evento no las trae.

    El dispatch del merge gobernado no las trae: no existen en su payload. Su
    ausencia se distingue de «el evento dice que no hubo force-push», y la
    salida lo dice en vez de callar un veredicto que no se midió (#1266). Un
    push que no las trae sigue fallando cerrado (HR-3).
    """
    forced, deleted = event.get("forced"), event.get("deleted")
    if forced is None and deleted is None:
        if is_dispatch(event):
            return None
        fail("event 'forced'/'deleted' flags are missing")
    if not isinstance(forced, bool) or not isinstance(deleted, bool):
        fail("event 'forced'/'deleted' flags are missing")
    return forced, deleted


def event_fields(event: dict) -> tuple[str, dict, tuple[bool, bool] | None, str, str]:
    """Rama, repositorio, banderas y rango auditado del evento.

    El rango se lee con el lector compartido `push_range` (#933): el dispatch
    lo lleva en `inputs` y el push en el nivel superior. Una forma dudosa es
    fallo cerrado (HR-3).
    """
    branch, repo = event_identity(event)
    flags = event_flags(event)
    before, after = push_range(event)
    return branch, repo, flags, before, after


def incident(event: dict, branch: str, flags: tuple[bool, bool], findings: list[str],
             rango: tuple[str, str]) -> tuple[str, str]:
    """Titulo y cuerpo de la incidencia; el rango llega ya resuelto (#1266)."""
    forced, deleted = flags
    before, after = rango
    head = event.get("head_commit") if isinstance(event.get("head_commit"), dict) else {}
    message = str(head.get("message", "")).splitlines()[0] if head.get("message") else "<none>"
    pusher = event.get("pusher") if isinstance(event.get("pusher"), dict) else {}
    title = f"incident: {'force-push' if forced else 'deletion'} on protected branch '{branch}'"
    body = ("Force-push guard (`check_force_push_guard`), rama protegida por contrato: "
            f"{branch}\n\npusher: {pusher.get('name', 'unknown')}\n"
            f"rango: {before}..{after}\n"
            f"forced: {forced} · deleted: {deleted}\n"
            f"head commit: {message}\n\n" + "\n".join(findings)
            + "\n\nLas ramas protegidas por contrato no se reescriben ni se borran; "
              "revierta con un revert normal y cierre este incidente con su post-mortem (HR-21).")
    return title, body


def audit(args: argparse.Namespace) -> int:
    event = load_json(args.event, "event")
    contract = load_json(args.contract, "contract")
    if not isinstance(event, dict) or not isinstance(contract, dict):
        fail("event and contract must be JSON objects")
    branch, repo, flags, before, after = event_fields(event)
    if branch not in protected_branches(contract, repo.get("default_branch")):
        print(f"SKIP: branch '{branch}' is not protected by contract; nothing to guard")
        return 0
    rango = f"{before} -> {after}"
    findings = []
    forced = deleted = False
    if flags is not None:
        forced, deleted = flags
        if forced:
            findings.append(f"force-push to protected branch '{branch}': {rango}")
        if deleted:
            findings.append(f"deletion of protected branch '{branch}': {rango}")
    if not findings:
        if flags is None:
            print(f"OK: dispatch on protected branch '{branch}' audited the range {rango}: "
                  "well-formed range, no force-push nor deletion reported for it; "
                  "the event carries no 'forced'/'deleted' flag")
        else:
            print(f"OK: push on protected branch '{branch}' is neither forced nor a deletion")
        return 0
    title, body = incident(event, branch, (forced, deleted), findings, (before, after))
    url = gh_open_incident(repo["full_name"], title, body)
    for finding in findings:
        print(f"  VIOLATION  {finding}")
    print(f"INCIDENT ISSUE: {url}")
    return 1


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="force-push / deletion guard vs the host contract")
    parser.add_argument("--event", required=True)
    parser.add_argument("--contract", required=True)
    args = parser.parse_args()
    try:
        code = audit(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
