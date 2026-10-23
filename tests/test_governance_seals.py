"""La fase 3 se mide dentro del pull request (HR-72, canon #1027).

`adoption check --phase 3` lee el árbol y el contrato: los sellos
`sha256` de `.governance-manifest.json` y el cableado de
`P23_required_jobs`. Antes de esta declaración la fase solo corría
programada (`operating-doc-drift.yml`, evento `schedule`), así que un
fichero gobernado podía llegar a `main` sin re-sellarse y el rojo
aparecía después, ya integrado.

La activación es por declaración explícita, como `agents_md` (HR-55) y
`docs_map` (HR-63): sin `governance_seals` el gate avisa a `stderr` y sale
0 para no romper la adopción previa. Estos tests miran la puerta desde las
dos caras: que la fase 3 deje de avisar (la declaración existe y está
cableada) y que el job que la mide tenga la forma que el canon exige, que
es lo único que la puerta no puede ver por sí sola.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CLI = ROOT / ".agents" / "skills" / "ci-pattern" / "assets" / "bin" / "ci-pattern"
WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
POLICY = ROOT / ".github" / "required-jobs-policy.json"
ADOPTION = ROOT / ".github" / "ci-pattern-adoption.json"
JOB = "governance-seals"


def _phase_three() -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(CLI), "adoption", "check", "--phase", "3", str(ROOT)],
        capture_output=True,
        text=True,
        cwd=ROOT,
    )


def _jobs() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))["jobs"]


def _job() -> dict:
    jobs = _jobs()
    assert JOB in jobs, (
        f"el job `{JOB}` no existe en {WORKFLOW.name}: el contrato lo declara y la "
        f"fase 3 lo exige, así que sin él la declaración queda sin cablear. "
        f"Jobs presentes: {sorted(jobs)}"
    )
    return jobs[JOB]


def test_phase_three_no_longer_reports_the_undeclared_class() -> None:
    """Con el job cableado, la fase 3 sale 0 y sin el aviso HR-72.

    El aviso y el rojo son dos caras de lo mismo: si la declaración falta,
    la fase no mide nada y sale 0 (fail-open); si la declaración está puesta
    y el job no, sale 1 nombrándolo. Aquí se exigen las dos cosas a la vez,
    que es el único estado en el que la fase 3 está de verdad cableada.
    """
    completed = _phase_three()

    assert "HR-72" not in completed.stderr, (
        "el contrato de adopción no declara `governance_seals`, así que la fase 3 "
        "sigue sin medirse en ningún pull request:\n" + completed.stderr.strip()
    )
    assert completed.returncode == 0, (
        f"la fase 3 debe salir 0 con el job cableado; rc={completed.returncode}\n"
        f"{completed.stderr.strip()}"
    )


def test_the_declared_job_is_the_one_the_canon_names() -> None:
    """La declaración usa el nombre que fija el canon, y ese job existe.

    `parameters.md` (nº 35) fija la convención `governance-seals`. Si el
    contrato declarara otro nombre, el gate lo mediría igual, pero el
    contrato se apartaría de la convención que leen los adopters.
    """
    contract = json.loads(ADOPTION.read_text(encoding="utf-8"))

    assert contract.get("governance_seals") == {"job": JOB}, (
        f"el contrato debe declarar `governance_seals: {{\"job\": \"{JOB}\"}}` "
        f"(convención del canon, `references/porting-guide.md`); "
        f"encontrado: {contract.get('governance_seals')!r}"
    )
    assert JOB in _jobs(), f"el job `{JOB}` que el contrato declara no existe en {WORKFLOW.name}"


def test_the_job_is_required_and_aggregated() -> None:
    """Sin entrada en `required_jobs` y en el `needs` del agregador, el job corre y nadie lo exige."""
    policy = json.loads(POLICY.read_text(encoding="utf-8"))
    required = _jobs()["required"]

    assert JOB in policy["required_jobs"], (
        f"`{JOB}` no figura en `required_jobs` de {POLICY.name}: en este repositorio "
        "el único check que el host exige es el agregador `required`, así que un job "
        "fuera de la política se ejecuta y no bloquea nada (HR-72)"
    )
    needs = required["needs"]
    needs = [needs] if isinstance(needs, str) else list(needs)
    assert JOB in needs, f"`{JOB}` no está en el `needs:` del agregador `required`: {needs}"


def test_the_job_publishes_its_check_only_from_pull_request() -> None:
    """HR-30: en `push`, `schedule` y `workflow_dispatch` no hay PR que evaluar.

    Publicar el nombre desde un evento sin PR daría un check que aparenta
    haber medido una fase que no midió. Por eso los tres eventos quedan en
    `accepted_skips` de la política y el nombre sale solo de `pull_request`.
    """
    condition = str(_job().get("if") or "")

    assert "pull_request" in condition, f"el job no corre en pull_request: {condition!r}"
    for event in ("push", "schedule", "workflow_dispatch"):
        assert event not in condition, (
            f"el job no debe correr en `{event}` (HR-30): {condition!r}"
        )


def test_the_job_invokes_the_cli_as_a_top_level_command() -> None:
    """HR-46: la CLI es el comando del paso, nunca el argumento de otro gate.

    `ci-pattern verify` ejecutaba `adoption check` en proceso y el ciclo
    superó 1.300 procesos de python y tumbó el host por OOM el 2026-10-05.
    El paso tiene que llamar a la CLI directamente.
    """
    commands = "\n".join(str(step.get("run") or "") for step in _job()["steps"])

    assert "adoption check --phase 3" in commands, (
        "el job no ejecuta `adoption check --phase 3`; sin eso el nombre del check "
        "publica un verde que no ha medido nada"
    )
    assert "verify" not in commands, (
        "el job no debe pasar por `ci-pattern verify` (HR-46): esa vía ejecuta "
        "`adoption check` en proceso y es el ciclo que produjo el OOM del 2026-10-05"
    )


def test_the_job_uses_no_secret_and_no_self_hosted_runner() -> None:
    """Es un job alcanzable desde un PR de un repositorio público.

    Los PR de fork no reciben secretos, así que un gate que solo puede
    medir con secreto se salta justo donde tiene que ser más duro. Y
    `scripts/check_workflows.py` prohíbe `self-hosted`: el runner es
    efímero y hospedado.
    """
    job = _job()
    rendered = yaml.safe_dump(job, allow_unicode=True)

    assert "self-hosted" not in rendered, (
        f"el job declara un runner propio: {job.get('runs-on')!r}. En un repositorio "
        "público un PR es entrada no confiable y `check_workflows.py` lo prohíbe"
    )
    assert "secrets." not in rendered, (
        "el job pide un secreto. Los PR de fork no lo reciben, así que el gate "
        "se saltaría en el evento que lo hace obligatorio"
    )
