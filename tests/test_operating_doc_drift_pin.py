"""El commit del catálogo de ``operating-doc-drift`` se lee del sello, no se escribe (#1157).

Antes el ``ref:`` del checkout del catálogo era un SHA escrito a mano, y cada
propagación de team-skills dejaba el pin atrás del ``.team-skills.yaml``. Ahora el
paso ``catalog-pin`` resuelve el commit desde el sello ``catalog @`` y el checkout
toma ese output. Estos tests fallan si el pin vuelve a escribirse a mano, y ejecutan
el paso de resolución de verdad contra manifiestos de mentira.
"""

from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

import yaml

WORKFLOWS_DIR = Path(__file__).resolve().parents[1] / ".github" / "workflows"


class _wy:
    """Lectura estructural mínima de un workflow (sin el helper de APAP)."""

    WORKFLOWS_DIR = WORKFLOWS_DIR

    @staticmethod
    def load(path: Path) -> dict:
        return yaml.safe_load(path.read_text(encoding="utf-8"))

    @staticmethod
    def job(doc: dict, job_id: str) -> dict:
        return doc["jobs"][job_id]

    @staticmethod
    def find_step(job: dict, name: str) -> dict:
        return next(s for s in job["steps"] if s.get("name") == name)


wy = _wy

WORKFLOW = wy.WORKFLOWS_DIR / "operating-doc-drift.yml"
CI_PATTERN_YAML = Path(__file__).resolve().parents[1] / "ci-pattern.yaml"
RESOLVER_STEP = "resolver el commit del catálogo desde su sello (#1157)"
CATALOG_REPO = "DysTelefonica/team-skills"
GOOD_COMMIT = "1234567890abcdef1234567890abcdef12345678"
SHORT_SEAL = "1234567"
SEAL = "# Auto-generado por propagate-team-skills.ps1 (catalog @ {commit})."

SHA40 = re.compile(r"(?<![0-9a-fA-F])[0-9a-f]{40}(?![0-9a-fA-F])")
#: El ref de una action en `uses:` es el único sha de 40 hex que un workflow puede
#: llevar escrito (HR-26): no es el pin del catálogo.
USES_REF = re.compile(r"\buses:\s*[A-Za-z0-9._/-]+@[0-9a-f]{40}")
#: Un pin literal que se inyecta a propósito para comprobar que el detector lo ve.
LITERAL_PIN = "bc915526af7920e459afd2532e93e96e0479c0e6"


def hand_written_shas(text: str) -> list[str]:
    """Los sha de 40 hex escritos a mano, fuera del ref de una action (`uses:`)."""
    found = []
    for n, line in enumerate(text.splitlines(), 1):
        for match in SHA40.finditer(USES_REF.sub("", line)):
            found.append(f"linea {n}: {match.group(0)}")
    return found


def catalog_checkout(doc: dict) -> dict:
    """El `with:` del checkout que trae el catálogo de team-skills."""
    job = wy.job(doc, "operating-doc-drift")
    for step in job["steps"]:
        with_ = step.get("with") if isinstance(step, dict) else None
        if isinstance(with_, dict) and with_.get("repository") == CATALOG_REPO:
            return with_
    raise AssertionError(f"ningún checkout trae {CATALOG_REPO}: el job no mide la deriva")


def resolver(doc: dict) -> dict:
    return wy.find_step(wy.job(doc, "operating-doc-drift"), RESOLVER_STEP)


class TestOperatingDocDriftPin:
    def test_el_checkout_del_catalogo_toma_su_ref_del_paso_de_sello(self):
        doc = wy.load(WORKFLOW)
        step = resolver(doc)
        assert step.get("id") == "catalog-pin", (
            "el paso que resuelve el sello necesita `id: catalog-pin`"
        )
        assert catalog_checkout(doc)["ref"] == "${{ steps.catalog-pin.outputs.commit }}"

    def test_el_paso_de_sello_corre_antes_que_el_checkout_del_catalogo(self):
        steps = wy.job(wy.load(WORKFLOW), "operating-doc-drift")["steps"]
        resolve_at = next(i for i, s in enumerate(steps) if s.get("name") == RESOLVER_STEP)
        checkout_at = next(
            i
            for i, s in enumerate(steps)
            if isinstance(s.get("with"), dict) and s["with"].get("repository") == CATALOG_REPO
        )
        assert resolve_at < checkout_at, (
            "el checkout del catálogo necesita el output del paso de sello"
        )

    def test_no_hay_sha_del_catalogo_escrito_en_el_workflow_ni_en_ci_pattern_yaml(self):
        assert hand_written_shas(WORKFLOW.read_text(encoding="utf-8")) == []
        assert hand_written_shas(CI_PATTERN_YAML.read_text(encoding="utf-8")) == []

    def test_el_detector_ve_un_ref_literal_inyectado(self):
        # Caso negativo: el checkout con el pin escrito a mano tiene que aparecer.
        injected = WORKFLOW.read_text(encoding="utf-8").replace(
            "ref: ${{ steps.catalog-pin.outputs.commit }}", f"ref: {LITERAL_PIN}", 1
        )
        assert injected != WORKFLOW.read_text(encoding="utf-8"), "el fixture no sustituyó el ref"
        assert hand_written_shas(injected), "el detector no ve un pin escrito a mano"

    # --- el paso se ejecuta de verdad sobre manifiestos de mentira ---

    def _run(self, tmp_path: Path, manifest: str | None, api: dict[str, str] | None = None):
        step = resolver(wy.load(WORKFLOW))
        env = {
            **os.environ,
            **{k: str(v) for k, v in step.get("env", {}).items() if "${{" not in str(v)},
            "CATALOG_TOKEN": "fixture",
            "GITHUB_OUTPUT": str(tmp_path / "github-output"),
        }
        (tmp_path / "github-output").write_text("", encoding="utf-8")
        if manifest is not None:
            (tmp_path / ".team-skills.yaml").write_text(manifest, encoding="utf-8")
        if api is not None:
            for rel, body in api.items():
                target = tmp_path / "api" / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text(body, encoding="utf-8")
            env["CATALOG_API"] = f"file://{tmp_path / 'api'}"
        script = tmp_path / "resolve.sh"
        script.write_text(str(step["run"]), encoding="utf-8")
        proc = subprocess.run(
            ["bash", str(script)],
            cwd=tmp_path,
            env=env,
            capture_output=True,
            text=True,
            timeout=60,
        )
        return proc, (tmp_path / "github-output").read_text(encoding="utf-8")

    def test_con_el_sello_completo_exporta_el_commit(self, tmp_path):
        proc, exported = self._run(
            tmp_path, f"{SEAL.format(commit=GOOD_COMMIT)}\nproject_type: web\n"
        )
        assert proc.returncode == 0, proc.stderr
        assert f"commit={GOOD_COMMIT}" in exported

    def test_con_el_sello_abreviado_lo_expande_la_api_del_catalogo(self, tmp_path):
        # El sello que hay hoy en el campo es `git rev-parse --short`.
        proc, exported = self._run(
            tmp_path,
            f"{SEAL.format(commit=SHORT_SEAL)}\nproject_type: web\n",
            api={f"repos/{CATALOG_REPO}/commits/{SHORT_SEAL}": f'{{"sha": "{GOOD_COMMIT}"}}'},
        )
        assert proc.returncode == 0, proc.stderr
        assert f"commit={GOOD_COMMIT}" in exported

    def test_sin_manifiesto_falla_nombrando_el_fichero(self, tmp_path):
        proc, exported = self._run(tmp_path, None)
        assert proc.returncode != 0, "sin .team-skills.yaml el job no puede salir en verde (HR-3)"
        assert ".team-skills.yaml" in proc.stderr
        assert "commit=" not in exported

    def test_sin_sello_falla_nombrando_el_sello(self, tmp_path):
        proc, exported = self._run(tmp_path, "project_type: web\ntiers:\n  - universal\n")
        assert proc.returncode != 0
        assert "catalog @" in proc.stderr
        assert "commit=" not in exported

    def test_un_prefijo_que_la_api_no_resuelve_falla_nombrando_el_sha(self, tmp_path):
        proc, exported = self._run(
            tmp_path, f"{SEAL.format(commit=SHORT_SEAL)}\nproject_type: web\n", api={}
        )
        assert proc.returncode != 0, "un sello que el catálogo no resuelve no se puede medir (HR-3)"
        assert SHORT_SEAL in proc.stderr
        assert "commit=" not in exported
