#!/usr/bin/env python3
# ci-pattern asset — merge-target gate (DysTelefonica/team-skills#420)
"""Gate de base de merge: una PR apunta a la rama por defecto, salvo cadena.

Hallazgo de auditoria (APAP_WEB, 2026-10-12): una PR apilada se mergeo en su
rama base en lugar de en `main` — ningun gate lo impidio, y en repos privados
Free (sin rulesets) no hay barrera del host. Este gate lo es:

- `base.ref` == rama por defecto -> OK (0).
- `base.ref` != defecto:
    - valido SOLO si el PR declara la etiqueta de cadena (`chain:partial`
      por defecto) Y su base es el `head.ref` de OTRA PR abierta (fichero
      de datos pasado por el workflow, mismo patron que el meta-gate HR con
      `--open-issues-file`); -> OK (0)
    - en cualquier otro caso -> hallazgo (1), nombrando la causa exacta.

Sin API: el workflow serializa `github.event` a un JSON y lista las cabeceras
de las PR abiertas con `gh pr list --json headRefName`. Cualquier lectura
imposible (payload ausente, JSON malformado, sin `pull_request`, sin base,
fichero de PRs abiertas ilegible o vacio) es entrada invalida: exit 2, nunca
un verde sobre evidencia no leida (HR-3). El gate no invoca la CLI del patron
como subproceso (#342).

Exit codes: 0 conforme · 1 hallazgos · 2 entrada invalida.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

DEFAULT_LABEL = "chain:partial"


class Invalid(Exception):
    """Entrada no interpretable: mapa a exit 2 (HR-3)."""


def load_event(path: str) -> dict:
    p = Path(path)
    if not p.is_file():
        raise Invalid(f"payload del evento ausente: {path}")
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise Invalid(f"payload del evento no es JSON valido: {exc}") from exc
    pr = data.get("pull_request")
    if not isinstance(pr, dict):
        raise Invalid("el payload no contiene el objeto `pull_request`.")
    base = pr.get("base") or {}
    base_ref = base.get("ref")
    if not isinstance(base_ref, str) or not base_ref.strip():
        raise Invalid("el payload no declara `pull_request.base.ref`.")
    labels = pr.get("labels")
    if not isinstance(labels, list):
        labels = []
    return {
        "base_ref": base_ref.strip(),
        "number": pr.get("number"),
        "labels": {str(item.get("name", "")) for item in labels if isinstance(item, dict)},
    }


def load_open_heads(path: str) -> set[str]:
    p = Path(path)
    if not p.is_file():
        raise Invalid(f"fichero de PRs abiertas ausente: {path}")
    raw = p.read_text(encoding="utf-8")
    if raw.startswith("["):  # variante JSON: gh pr list --json headRefName
        try:
            rows = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise Invalid(f"lista de PRs abiertas ilegible: {exc}") from exc
        heads = {str(r.get("headRefName", "")) for r in rows if isinstance(r, dict)}
    else:  # variante linea a linea (salida de --jq)
        heads = {line.strip() for line in raw.splitlines() if line.strip()}
    heads.discard("")
    if not heads:
        raise Invalid("la lista de PRs abiertas salio vacia: es un fallo de la "
                      "API o del paso previo, no cero PRs abiertas (HR-3).")
    return heads


def evaluate(pr: dict, open_heads: set[str], default_branch: str,
             chain_label: str) -> list[str]:
    if pr["base_ref"] == default_branch:
        return []
    if chain_label not in pr["labels"]:
        return [
            f"HR-59: la PR #{pr['number']} tiene base '{pr['base_ref']}' distinta "
            f"de la rama por defecto '{default_branch}' sin la etiqueta de cadena "
            f"'{chain_label}'. Base incorrecta: reabrir contra '{default_branch}' "
            "o corregir la etiqueta del intermedio de cadena. (#420)"
        ]
    if pr["base_ref"] not in open_heads:
        return [
            f"HR-59: la PR #{pr['number']} declara cadena ({chain_label}) pero su "
            f"base '{pr['base_ref']}' no es el head de ninguna PR abierta: la "
            "cadena apunta a una rama huesped o ya cerrada. (#420)"
        ]
    return []


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv=None) -> int:
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="merge-target gate (#420)")
    ap.add_argument("--event", required=True,
                    help="JSON con el payload github.event (objeto pull_request)")
    ap.add_argument("--open-prs", required=True,
                    help="fichero con las headRefName de las PR abiertas "
                         "(JSON o una por linea)")
    ap.add_argument("--default-branch", default="main",
                    help="rama por defecto del consumer (default: main)")
    ap.add_argument("--chain-label", default=DEFAULT_LABEL,
                    help="etiqueta de intermedio de cadena (default: chain:partial)")
    args = ap.parse_args(argv)

    try:
        pr = load_event(args.event)
        open_heads = load_open_heads(args.open_prs)
    except Invalid as exc:
        print(f"ABORT: {exc}", file=sys.stderr)
        return 2

    findings = evaluate(pr, open_heads, args.default_branch, args.chain_label)
    for f in findings:
        print(f, file=sys.stderr)
    if findings:
        return 1
    if pr["base_ref"] != args.default_branch:
        print(f"OK: cadena valida — base '{pr['base_ref']}' es el head de otra "
              f"PR abierta ({args.chain_label}).")
    else:
        print(f"OK: base = rama por defecto '{args.default_branch}'.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
