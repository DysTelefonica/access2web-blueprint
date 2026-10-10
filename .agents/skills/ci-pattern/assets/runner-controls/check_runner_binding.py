#!/usr/bin/env python3
# ci-pattern asset — runner binding (DysTelefonica/team-skills#303)
"""Binding ejecutable entre la declaracion runner-enforced y los workflows.

Una regla `runner-enforced` declarada en el contrato del host sin un
workflow que la aplique es autodeclaracion, no gate. Este asset verifica
el binding en el repo destino:

- cada AREA del contrato con alguna regla `runner-enforced` exige un job,
  en algun workflow del directorio, cuyo `runs-on` incluya la etiqueta
  `runner.label` y que invoque al asset que cubre el area — `labels` y
  `rulesets` -> `check_governed_merge.py`; `required_checks` ->
  `check_push_compliance.py` o `check_governed_merge.py`; `protection` ->
  `check_force_push_guard.py`;
- y los tres controles compensatorios (detective, merge gobernado, guard)
  deben estar cableados en jobs de esa etiqueta, porque un merge fuera del
  camino gobernado solo lo marca el detective.

Sin cablear: exit 1 nombrando el asset y la etiqueta. Entrada insalvable
(contrato malformado, `runner.label` ausente con reglas runner-enforced,
directorio de workflows inexistente, workflow sin `jobs:`): exit 2 (HR-3).
Solo stdlib, sin red; el escaneo es a nivel de job, no de fichero, para
no contar un asset ejecutado en un runner ajeno.
"""
import argparse
import json
import re
import sys
from pathlib import Path

from check_push_compliance import FailClosed

DETECTIVE = "check_push_compliance.py"
MERGE = "check_governed_merge.py"
GUARD = "check_force_push_guard.py"
ALL_ASSETS = (DETECTIVE, MERGE, GUARD)
AREA_ASSETS = {"labels": {MERGE}, "rulesets": {MERGE},
               "required_checks": {DETECTIVE, MERGE}, "protection": {GUARD}}
JOB_KEY = re.compile(r"^  ([A-Za-z0-9_-]+):\s*$")
RUNS_ON = re.compile(r"^\s*runs-on:\s*(.+)$")


def load_json(path: str, what: str) -> object:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except OSError as exc:
        raise FailClosed(f"unreadable {what} '{path}': {exc}") from exc
    except json.JSONDecodeError as exc:
        raise FailClosed(f"{what} '{path}' is not valid JSON: {exc}") from exc


def declared_areas(contract: dict) -> tuple[dict, str | None]:
    """Areas con reglas runner-enforced y la etiqueta del runner."""
    sections = {"labels": [e.get("class") for e in contract.get("labels", []) if isinstance(e, dict)],
                "required_checks": [e.get("class") for e in contract.get("required_checks", []) if isinstance(e, dict)],
                "protection": [v.get("class") for v in contract.get("protection", {}).values()
                               if isinstance(v, dict)],
                "rulesets": [e.get("class") for e in contract.get("rulesets", []) if isinstance(e, dict)]}
    areas = {name for name, classes in sections.items() if "runner-enforced" in classes}
    runner = contract.get("runner", {})
    if not isinstance(runner, dict) or (areas and (
            not isinstance(runner.get("label"), str) or not runner["label"])):
        raise FailClosed("runner-enforced rules require contract key 'runner' with a non-empty 'label'")
    label = runner.get("label") if isinstance(runner.get("label"), str) else None
    return areas, label


def job_blocks(text: str, name: str) -> list[str]:
    """Bloques de job (indentacion 2) de la seccion `jobs:`."""
    if not re.search(r"^jobs:\s*$", text, re.M):
        raise FailClosed(f"workflow '{name}' has no 'jobs:' section: malformed input")
    blocks: list[str] = []
    current: str | None = None
    for line in text.splitlines():
        if re.match(r"^jobs:\s*$", line):
            continue
        if current is not None and line.strip() and not line.startswith("  "):
            blocks.append(current)
            current = None
        match = JOB_KEY.match(line)
        if match:
            if current is not None:
                blocks.append(current)
            current = line + "\n"
        elif current is not None:
            current += line + "\n"
    if current is not None:
        blocks.append(current)
    return blocks


def block_tokens(block: str) -> set[str]:
    """Tokens declarados en los `runs-on:` de un bloque de job."""
    tokens: set[str] = set()
    for line in block.splitlines():
        match = RUNS_ON.match(line)
        if match:
            tokens |= {t.strip().strip("\"'")
                       for t in re.split(r"[\s,\[\]]+", match.group(1)) if t.strip()}
    return tokens


def wired_files(wdir: Path, label: str, asset: str) -> list[Path]:
    """Workflows con al menos un job cuyo runs-on incluye `label` y cuyo
    bloque invoca `asset`. Unica fuente de la deteccion workflow→asset
    (HR-52); el documento operativo (#351) la reutiliza en el mismo
    proceso en vez de adivinar por nombre de fichero."""
    found: list[Path] = []
    for path in sorted(list(wdir.glob("*.yml")) + list(wdir.glob("*.yaml"))):
        text = path.read_text(encoding="utf-8")
        for block in job_blocks(text, path.name):
            if label in block_tokens(block) and asset in block:
                found.append(path)
                break
    return found


def wired_assets(wdir: Path, label: str) -> set[str]:
    """Assets invocados por jobs cuyo runs-on incluye la etiqueta."""
    return {asset for asset in ALL_ASSETS if wired_files(wdir, label, asset)}


def main() -> int:
    parser = argparse.ArgumentParser(description="runner-enforced declaration vs workflow wiring")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--workflows", required=True, help="consumer .github/workflows directory")
    args = parser.parse_args()
    try:
        contract = load_json(args.contract, "contract")
        if not isinstance(contract, dict):
            raise FailClosed("contract must be a JSON object")
        areas, label = declared_areas(contract)
        wdir = Path(args.workflows)
        if not wdir.is_dir():
            raise FailClosed(f"workflows directory '{wdir}' does not exist")
        if not label:
            print("SKIP: no runner-enforced rules declared; nothing to bind")
            return 0
        wired = wired_assets(wdir, label)
        findings = []
        for area in sorted(areas):
            coverage = AREA_ASSETS[area]
            if not (coverage & wired):
                findings.append(f"area '{area}' is declared runner-enforced but no workflow on runner "
                                f"'{label}' invokes any of {sorted(coverage)}")
        for asset in ALL_ASSETS:
            if asset not in wired:
                findings.append(f"compensating asset '{asset}' is not wired in any workflow "
                                f"on runner '{label}'")
        if findings:
            for finding in findings:
                print(f"  UNBOUND  {finding}")
            print(f"FAIL: {len(findings)} unbound runner-enforced declaration(s)")
            return 1
        print("RUNNER BINDING REPORT")
        for area in sorted(areas):
            print(f"  ok  area '{area}' enforced by {sorted(AREA_ASSETS[area] & wired)} on '{label}'")
        for asset in ALL_ASSETS:
            print(f"  ok  compensating asset {asset} wired on '{label}'")
        print(f"BINDING OK: runner-enforced declarations match the workflows on '{label}'")
        return 0
    except FailClosed as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
