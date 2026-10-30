#!/usr/bin/env python3
# ci-pattern asset tests — gate de imágenes de CI fijadas por digest (team-skills#1264).
"""check_ci_images_pinned.py observado como proceso, sobre repos mínimos en tmp_path."""

from __future__ import annotations

import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

GATE = Path(__file__).resolve().parents[1] / "check_ci_images_pinned.py"
DIG = "@sha256:" + "a" * 64
ON = "P56_ci_images_gate: on\n"
CONTAINER = "container:\n  image: alpine:3.20\n"
ECHO = "steps:\n  - run: echo hi\n"
RUN = "steps:\n  - run: |\n"


def run(tmp_path: Path, params, wf=None):
    """Repo mínimo: contrato (opcional) y el workflow `ci.yml` tal cual."""
    assert GATE.is_file(), f"falta el gate: {GATE}"
    if params is not None:
        (tmp_path / "ci-pattern.yaml").write_text(params, encoding="utf-8")
    if wf is not None:
        (tmp_path / ".github" / "workflows").mkdir(parents=True, exist_ok=True)
        (tmp_path / ".github" / "workflows" / "ci.yml").write_text(wf, encoding="utf-8")
    return subprocess.run(
        [sys.executable, str(GATE), str(tmp_path)], capture_output=True, text=True, timeout=60
    )


def job(body: str) -> str:
    """Workflow de un job `build` con `body` bajo él."""
    head = "on: push\njobs:\n  build:\n    runs-on: ubuntu-latest\n"
    return head + textwrap.indent(textwrap.dedent(body), "    ")


@pytest.mark.parametrize("params", ["P51_runners_profile: hosted\n", "P56_ci_images_gate: off\n"])
def test_p56_ausente_u_off_es_inerte(tmp_path, params):
    res = run(tmp_path, params, job(CONTAINER))
    assert res.returncode == 0 and res.stdout.startswith("PINES: off"), res.stderr


def test_entradas_invalidas_salen_2_nombrando_el_dato(tmp_path):
    """P56 inválido y workflow no parseable: exit 2 nombrando lo que falla."""
    res = run(tmp_path, "P56_ci_images_gate: maybe\n")
    assert res.returncode == 2 and "P56_ci_images_gate" in res.stderr
    res = run(tmp_path, ON, "on: push\n")
    assert res.returncode == 2 and "ci.yml" in res.stderr


@pytest.mark.parametrize("params", [None, ON + "\tP57_ci_image_exceptions: {}\n"])
def test_ci_pattern_yaml_ausente_o_ilegible_sale_2(tmp_path, params):
    res = run(tmp_path, params, job(ECHO))
    assert res.returncode == 2 and "ci-pattern.yaml" in res.stderr


def test_sin_directorio_de_workflows_es_limpio(tmp_path):
    res = run(tmp_path, ON)
    assert res.returncode == 0 and "workflows: 0" in res.stdout, res.stderr


BUILDX = "steps:\n  - uses: docker/setup-buildx-action@v3\n    with:\n"
SURFACES = [
    (CONTAINER, "container.image -> alpine:3.20"),
    ("services:\n  db:\n    image: postgres:16\n", "services.db.image -> postgres:16"),
    (BUILDX + "      driver-opts: image=moby/bk:v1\n", "steps[0].with.driver-opts -> moby/bk:v1"),
    (RUN + "      docker run \\\n        alpine:3.20 true\n", "steps[0].run -> alpine:3.20"),
    ("steps:\n  - run: docker pull alpine:3.20\n", "steps[0].run -> alpine:3.20"),
    (RUN + "      docker build -t app:ci .\n      docker pull img:1\n", "steps[0].run -> img:1"),
    ("steps:\n  - uses: docker://alpine:3.20\n", "steps[0].uses -> alpine:3.20"),
]


@pytest.mark.parametrize("body, finding", SURFACES)
def test_superficie_sin_digest_es_hallazgo(tmp_path, body, finding):
    res = run(tmp_path, ON, job(body))
    assert res.returncode == 1
    assert f"IMAGEN-SIN-DIGEST ci.yml:build:{finding}" in res.stdout.splitlines()


def test_etiqueta_que_construye_el_workflow_no_es_hallazgo(tmp_path):
    """Una etiqueta local no se descarga del registro: ni en build ni en run."""
    body = (
        RUN + "      docker build -t app:ci .\n      docker build -t app:ci-2 app:ci\n"
        "      docker build -t app:x ./ctx\n"
        "      docker build --tag=app:eq .\n      docker run app:eq true\n"
        "      docker build --tag test-app:pr .\n"
        "      docker run -d --name test-app --network host -e A=1 test-app:pr\n"
    )
    res = run(tmp_path, ON, job(body))
    assert res.returncode == 0 and "IMAGEN-SIN-DIGEST" not in res.stdout, res.stdout


def test_pull_de_etiqueta_construida_por_el_workflow_es_hallazgo(tmp_path):
    """Extraer del registro sí es una descarga: la exención no aplica al pull."""
    body = RUN + "      docker build --tag test-app:pr .\n      docker pull test-app:pr\n"
    res = run(tmp_path, ON, job(body))
    assert res.returncode == 1
    assert "IMAGEN-SIN-DIGEST ci.yml:build:steps[0].run -> test-app:pr" in res.stdout


PINNED = r"""
container:
  image: alpine:3.20DIG
services:
  db:
    image: postgres:16DIG
steps:
  - uses: docker/setup-buildx-action@v3
    with:
      driver-opts: image=moby/buildkit:v0.13.2DIG
  - run: |
      docker run \
        alpine:3.20DIG true
  - run: docker pull alpine:3.20DIG
  - uses: docker://alpine:3.20DIG
""".replace("DIG", DIG)


def test_todas_fijadas_es_limpio(tmp_path):
    """Fijadas por digest, `docker://` con digest y flags con valor: limpio."""
    res = run(tmp_path, ON, job(PINNED))
    assert res.returncode == 0 and "hallazgos: 0" in res.stdout, res.stdout
    body = (
        RUN
        + '      docker run --rm -v "${{ github.workspace }}:/work" alpine:3.20'
        + DIG
        + " true\n"
    )
    res = run(tmp_path, ON, job(body))
    assert res.returncode == 0 and "hallazgos: 0" in res.stdout, res.stdout
    res = run(tmp_path, ON, job(f"steps:\n  - uses: docker://alpine{DIG}\n"))
    assert res.returncode == 0 and "hallazgos: 0" in res.stdout, res.stdout
    body = (
        RUN + f"      docker run --gpus all --memory 2g --cap-add NET_ADMIN alpine:3.20{DIG} true\n"
    )
    res = run(tmp_path, ON, job(body))
    assert res.returncode == 0 and "hallazgos: 0" in res.stdout, res.stdout


def test_excepciones_declaradas(tmp_path):
    """Con motivo cubre la referencia; sin motivo es hallazgo; la que no casa, obsoleta."""
    res = run(
        tmp_path,
        ON + 'P57_ci_image_exceptions:\n  "alpine:3.20": imagen de tercero\n',
        job(CONTAINER),
    )
    assert res.returncode == 0 and "EXCEPCION" not in res.stdout, res.stdout
    res = run(tmp_path, ON + 'P57_ci_image_exceptions:\n  "alpine:3.20": ""\n', job(CONTAINER))
    assert res.returncode == 1
    assert "EXCEPCION-SIN-MOTIVO alpine:3.20" in res.stdout.splitlines()
    res = run(tmp_path, ON + 'P57_ci_image_exceptions:\n  "ghcr.io/x/old:1": motivo\n', job(ECHO))
    assert res.returncode == 1
    assert "EXCEPCION-OBSOLETA ghcr.io/x/old:1" in res.stdout.splitlines()


def test_referencia_no_literal_es_no_verificada(tmp_path):
    res = run(tmp_path, ON, job("container:\n  image: ${{ matrix.image }}\n"))
    assert res.returncode == 0, res.stdout
    assert "NO-VERIFICADA ci.yml:build:container.image -> ${{ matrix.image }}" in res.stdout
