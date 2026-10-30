"""HR-65: la declaración de `docs_commands` está atada al job que la mide.

Desde #1020 la fuente única es la clave ``docs_commands`` del contrato de
adopción (``.github/ci-pattern-adoption.json``); el fichero suelto
``docs-commands.json`` quedó obsoleto y ``run_docs_commands.py --config`` solo
avisa por ``stderr``. Con una sola declaración, el modo en que el contrato puede
volver a mentir es otro: prometer un job que no existe, un job que no ejecuta el
gate, o un job que nada espera por él. La fase 5 comprueba que el job declarado
existe y es requerido; nada ata que ese job mida lo que el contrato declara.

Este test ata esa cadena, por el mismo motivo que ``test_check_runner_profile``
ata el pin del catálogo entre ``ci-pattern.yaml`` y ``operating-doc-drift.yml``
(#1319).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml

REPO = Path(__file__).resolve().parent.parent
CONTRACT = REPO / ".github" / "ci-pattern-adoption.json"
WORKFLOW = REPO / ".github" / "workflows" / "ci.yml"
POLICY = REPO / ".github" / "required-jobs-policy.json"
# Obsoleto desde #1020; la fuente única es `docs_commands` en el contrato.
CONFIG_OBSOLETO = REPO / "docs-commands.json"

# Claves que el gate necesita; el resto es opcional y puede crecer sin romper la
# atada.
REQUIRED_KEYS = ("cli_name", "cli_command", "executed_verbs", "dry_run_flag", "job")


@pytest.fixture(scope="module")
def declared() -> dict:
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    return contract["docs_commands"]


@pytest.fixture(scope="module")
def workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def test_the_contract_declares_docs_commands() -> None:
    """Sin la declaración, `adoption check --phase 5` solo avisa por stderr.

    Falla si alguien la quita creyendo que es opcional: quitarla devuelve HR-65
    al estado de aviso que este cambio cierra.
    """
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    assert "docs_commands" in contract, (
        "el contrato no declara docs_commands: HR-65 vuelve a ser un aviso por "
        "stderr que no bloquea nada"
    )


@pytest.mark.parametrize("key", REQUIRED_KEYS)
def test_the_declaration_has_the_keys_the_gate_needs(declared: dict, key: str) -> None:
    assert key in declared, (
        f"docs_commands no declara {key!r}: run_docs_commands.py y la fase 5 la "
        "usan, y sin ella el job no puede probar nada"
    )


def test_the_declaration_declares_only_schema_keys(declared: dict) -> None:
    """``additionalProperties: false`` es la regla del contrato.

    Una clave de más la rechazaría el gate en runtime, no la fase 5.
    """
    schema = json.loads(
        (
            REPO / ".agents/skills/ci-pattern/assets/adoption/adoption-contract.schema.json"
        ).read_text(encoding="utf-8")
    )
    known = set(schema["properties"]["docs_commands"]["properties"])
    unknown = sorted(set(declared) - known)
    assert not unknown, f"claves que el esquema no declara: {unknown}"


def test_the_obsolete_config_does_not_come_back() -> None:
    """La fuente única es el contrato; un segundo fichero vuelve a mentir.

    `docs-commands.json` en la raíz era la doble declaración anterior a #1020.
    Si vuelve, el contrato y el job pueden divergir en silencio otra vez, y el
    runner solo lo avisa por stderr.
    """
    assert not CONFIG_OBSOLETO.exists(), (
        "docs-commands.json está obsoleto desde #1020: la fuente única es "
        "docs_commands en .github/ci-pattern-adoption.json. Borra el fichero en "
        "vez de mantener dos declaraciones."
    )


def test_the_declared_job_exists_and_runs_the_gate(declared: dict, workflow: dict) -> None:
    """El job que el contrato nombra tiene que ser el que mide los comandos."""
    job_name = declared["job"]
    jobs = workflow["jobs"]
    assert job_name in jobs, (
        f"el contrato declara el job {job_name!r} y .github/workflows/ci.yml no "
        f"lo define; los jobs son {sorted(jobs)}"
    )
    runs = "\n".join(str(step.get("run", "")) for step in jobs[job_name].get("steps", []))
    assert "run_docs_commands.py" in runs, (
        f"el job {job_name!r} no ejecuta run_docs_commands.py: el gate corre, "
        "pero nada mide los comandos documentados"
    )


def test_the_job_is_required_in_the_policy(declared: dict) -> None:
    """El aviso por stderr sale aunque el job no bloquee; esto lo cierra."""
    policy = json.loads(POLICY.read_text(encoding="utf-8"))
    assert declared["job"] in policy["required_jobs"], (
        f"{declared['job']} no está en required_jobs: el gate corre pero nada "
        "espera por él y el mapa puede volver a mentir sin que nadie lo vea"
    )


def test_the_aggregator_waits_for_the_job(declared: dict, workflow: dict) -> None:
    """Estar requerido no basta si el agregador no lo espera.

    Sin el `needs`, el agregador puede terminar en verde antes de que el job
    mida nada, y el rojo del gate llegaría después del merge.
    """
    job_name = declared["job"]
    needs = workflow["jobs"]["required"]["needs"]
    assert job_name in needs, f"required no espera por {job_name!r}: needs = {needs}"
