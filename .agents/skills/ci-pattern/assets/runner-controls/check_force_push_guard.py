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
"""
import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

from check_push_compliance import FailClosed, gh_open_incident, protected_branches

GH_TIMEOUT = 30
SHA_RE = re.compile(r"^[0-9a-f]{40}$")


def load_json(path: str, what: str) -> object:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except OSError as exc:
        raise FailClosed(f"unreadable {what} '{path}': {exc}") from exc
    except json.JSONDecodeError as exc:
        raise FailClosed(f"{what} '{path}' is not valid JSON: {exc}") from exc


def main() -> int:
    parser = argparse.ArgumentParser(description="force-push / deletion guard vs the host contract")
    parser.add_argument("--event", required=True)
    parser.add_argument("--contract", required=True)
    args = parser.parse_args()
    try:
        event = load_json(args.event, "event")
        contract = load_json(args.contract, "contract")
        if not isinstance(event, dict) or not isinstance(contract, dict):
            raise FailClosed("event and contract must be JSON objects")
        ref = event.get("ref")
        if not isinstance(ref, str) or not ref.startswith("refs/heads/"):
            raise FailClosed("event has no 'refs/heads/' ref to audit")
        repo = event.get("repository")
        if not isinstance(repo, dict) or not isinstance(repo.get("full_name"), str) \
                or repo["full_name"].count("/") != 1:
            raise FailClosed("event has no OWNER/REPO 'repository.full_name'")
        forced, deleted = event.get("forced"), event.get("deleted")
        if not isinstance(forced, bool) or not isinstance(deleted, bool):
            raise FailClosed("event 'forced'/'deleted' flags are missing")
        before, after = event.get("before"), event.get("after")
        if not isinstance(before, str) or not SHA_RE.match(before) or not isinstance(after, str) \
                or not SHA_RE.match(after):
            raise FailClosed("event 'before'/'after' are not 40-char commit ids")
        branch = ref[len("refs/heads/"):]
        if branch not in protected_branches(contract, repo.get("default_branch")):
            print(f"SKIP: branch '{branch}' is not protected by contract; nothing to guard")
            return 0
        findings = []
        if forced:
            findings.append(f"force-push to protected branch '{branch}': {before} -> {after}")
        if deleted:
            findings.append(f"deletion of protected branch '{branch}': {before} -> {after}")
        if not findings:
            print(f"OK: push on protected branch '{branch}' is neither forced nor a deletion")
            return 0
        head = event.get("head_commit") if isinstance(event.get("head_commit"), dict) else {}
        message = str(head.get("message", "")).splitlines()[0] if head.get("message") else "<none>"
        pusher = event.get("pusher") if isinstance(event.get("pusher"), dict) else {}
        title = f"incident: {'force-push' if forced else 'deletion'} on protected branch '{branch}'"
        body = ("Force-push guard (`check_force_push_guard`), rama protegida por contrato: "
                f"{branch}\n\npusher: {pusher.get('name', 'unknown')}\n"
                f"rango: {before}..{after}\nforced: {forced} · deleted: {deleted}\n"
                f"head commit: {message}\n\n" + "\n".join(findings)
                + "\n\nLas ramas protegidas por contrato no se reescriben ni se borran; "
                  "revierta con un revert normal y cierre este incidente con su post-mortem (HR-21).")
        url = gh_open_incident(repo["full_name"], title, body)
        for finding in findings:
            print(f"  VIOLATION  {finding}")
        print(f"INCIDENT ISSUE: {url}")
        return 1
    except FailClosed as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
