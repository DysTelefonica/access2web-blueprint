#!/usr/bin/env python3
"""Suite del gate de paridad local/CI: árboles falsos en tmp (positivo,
negativo por dirección, abreviado/bloque e indentación) y dogfooding."""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_local_ci_parity.py"
REAL_POLICY = Path(__file__).resolve().parent / "fixtures" / "policy-team-skills.json"
REPO_ROOT = Path(__file__).resolve().parents[6]

WORKFLOW = """jobs:
  {job}:
    steps:
{steps}
"""

ENTRYPOINT = """#!/usr/bin/env bash
suites=(
  testing/suites/a/test-a.sh
  testing/suites/b/test-b.sh
)
for t in "${{suites[@]}}"; do
  if bash "$t"; then
    echo PASS
  fi
done
for t in personal/ardelperal/ci-pattern/assets/*/tests/test_*.py; do
  if python3 "$t"; then
    echo PASS
  fi
done
"""

CI_STEPS = """      - name: Suites desde el script único
        run: bash testing/run-unit-suites.sh
      - name: Gate CI-only sin declarar
        run: python3 personal/ardelperal/ci-pattern/assets/x/check_x.py --arg 1
"""

EXCL = [{"command": "python3 personal/ardelperal/ci-pattern/assets/x/check_x.py*", "side": "ci",
         "reason": "gate CI-only: requiere secretos del runner"}]

# Caso #917: el entrypoint local ejecuta `docker info` y la política lo marca
# como exclusión que necesita Docker. El gate solo debe aceptarlo si el
# entorno local declara la capacidad.
DOCKER_ENTRYPOINT = ENTRYPOINT + "docker info\n"
DOCKER_EXCL = [{"command": "docker info", "side": "any",
                "reason": "comprobación de Docker del entrypoint (#917)",
                "requires": ["docker"]}]


def build_repo(tmp, workflow=None, entrypoint=None):
    repo = Path(tmp) / "repo"
    wf = repo / ".github" / "workflows"
    wf.mkdir(parents=True, exist_ok=True)
    (wf / "tests.yml").write_text(workflow if workflow is not None
                                  else WORKFLOW.format(job="unit", steps=CI_STEPS), encoding="utf-8")
    ep = repo / "testing" / "run-unit-suites.sh"
    ep.parent.mkdir(parents=True, exist_ok=True)
    ep.write_text((entrypoint if entrypoint is not None else ENTRYPOINT).format(), encoding="utf-8")
    for rel in ("testing/suites/a/test-a.sh", "testing/suites/b/test-b.sh",
                "cmd/a.sh", "cmd/b.sh", "cmd/c.sh"):
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("#!/usr/bin/env bash\nexit 0\n", encoding="utf-8")
    return repo


def build_policy(tmp, exclusions, repo, job="unit", capabilities=None):
    policy = {"ci": {"workflow": ".github/workflows/tests.yml", "job": job},
              "local": {"entrypoint": "testing/run-unit-suites.sh"},
              "exclusions": exclusions}
    if capabilities is not None:
        policy["local"]["capabilities"] = capabilities
    path = Path(tmp) / "policy.json"
    path.write_text(json.dumps(policy), encoding="utf-8")
    return path


def run_gate(repo, policy):
    return subprocess.run([sys.executable, str(GATE), "--repo", str(repo), "--policy", str(policy)],
                          capture_output=True, text=True, check=False)


class LocalCiParityTests(unittest.TestCase):

    def test_parity_with_declared_exclusions_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp), build_policy(tmp, EXCL, tmp))
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("LOCAL/CI PARITY OK", proc.stdout)
            self.assertIn("bash testing/suites/a/test-a.sh", proc.stdout)

    def test_malformed_policy_fails_closed(self):
        """Política ilegible: exit 2 (HR-3)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp)
            path = Path(tmp) / "policy.json"
            path.write_text("{not json", encoding="utf-8")
            proc = run_gate(repo, path)
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("fail-closed", proc.stderr)

    def test_undeclared_ci_command_is_a_finding(self):
        """Gate del CI sin entrada local ni exclusión → hallazgo."""
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp), build_policy(tmp, [], tmp))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("comando del CI sin entrada local", proc.stdout)
            self.assertIn("check_x.py", proc.stdout)

    def test_flag_difference_between_local_and_ci_is_a_finding(self):
        """Friccion F1 (#396): dos lados ejecutando el mismo comando
        con distinto flag deben producir un hallazgo que nombre AMBAS
        formas exactas. Caso de referencia (defecto medido): el lint
        de ardelperal/APAP_WEB corria 'python3 scripts/check_alantyle.py
        --informational' en CI y 'python3 scripts/check_alantyle.py'
        local; con la _signature() colapsada a dos tokens, ambos lados
        daban 'python3 scripts/check_alantyle.py' y el gate daba falso
        verde sobre una suite RED (17 violaciones) detectadas fuera de
        la paridad. Tras el fix, los flags son visibles y el gate falla
        (HR-4 enmendada por #396)."""
        wf = ("jobs:\n"
              "  unit:\n"
              "    steps:\n"
              "      - run: bash testing/run-unit-suites.sh\n"
              "      - run: python3 scripts/check_alantyle.py --informational\n")
        ep = ('#!/usr/bin/env bash\n'
              'python3 scripts/check_alantyle.py\n')
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp, workflow=wf, entrypoint=ep)
            proc = run_gate(repo, build_policy(tmp, [], repo))
            self.assertEqual(proc.returncode, 1,
                             "el gate debe fallar (no dar falso verde) cuando "
                             "el CI corre el comando con --informational y el "
                             "local produce la forma sin flag "
                             "(friccion F1 #396):\n" + proc.stdout + proc.stderr)
            # El hallazgo debe nombrar ambas formas exactas del comando
            # para que el autor vea la diferencia de flag, no solo el lado
            # que difiere.
            self.assertIn("python3 scripts/check_alantyle.py --informational",
                          proc.stdout,
                          "el hallazgo debe nombrar la forma CI con flag")
            self.assertIn("python3 scripts/check_alantyle.py",
                          proc.stdout,
                          "el hallazgo debe nombrar la forma local sin flag")

    def test_undeclared_local_command_is_a_finding(self):
        """CI que NO invoca el entrypoint: comandos locales sin
        correspondencia (dirección 'viceversa' del criterio)."""
        with tempfile.TemporaryDirectory() as tmp:
            steps = "      - name: a\n        run: bash testing/suites/a/test-a.sh\n"
            repo = build_repo(tmp, workflow=WORKFLOW.format(job="unit", steps=steps))
            proc = run_gate(repo, build_policy(tmp, [], repo))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("comando local sin correspondencia", proc.stdout)
            self.assertIn("test-b.sh", proc.stdout)

    def test_relative_order_divergence_is_a_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            steps = ("      - name: b\n        run: bash testing/suites/b/test-b.sh\n"
                     "      - name: a\n        run: bash testing/suites/a/test-a.sh\n")
            repo = build_repo(tmp, workflow=WORKFLOW.format(job="unit", steps=steps))
            proc = run_gate(repo, build_policy(tmp, [], repo))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("orden relativo", proc.stdout)

    def test_local_entrypoint_running_command_that_needs_undeclared_capability_is_a_finding(self):
        """Caso que hoy falla (#917): el entrypoint local ejecuta un comando
        que necesita Docker y el entorno local no lo declara. Hoy el gate
        da verde porque no conoce capacidades; debe dar hallazgo."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp, entrypoint=DOCKER_ENTRYPOINT)
            proc = run_gate(repo, build_policy(tmp, DOCKER_EXCL + EXCL, repo))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("requiere ['docker']", proc.stdout)
            self.assertIn("docker info", proc.stdout)

    def test_declared_local_capability_lets_the_entrypoint_run_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp, entrypoint=DOCKER_ENTRYPOINT)
            proc = run_gate(repo, build_policy(tmp, DOCKER_EXCL + EXCL, repo,
                                               capabilities=["docker"]))
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("LOCAL/CI PARITY OK", proc.stdout)

    def test_capability_needed_only_by_ci_only_command_is_not_a_finding(self):
        """Un comando CI-only que necesita Docker no lo ejecuta el entrypoint:
        no exige capacidad local (solo motivo, como cualquier exclusión)."""
        excl = [{"command": "python3 personal/ardelperal/ci-pattern/assets/x/check_x.py*",
                 "side": "ci", "reason": "gate CI-only: requiere Docker en el runner",
                 "requires": ["docker"]}]
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp), build_policy(tmp, excl, tmp))
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_unknown_capability_fails_closed(self):
        """Vocabulario cerrado: una capacidad fuera de la lista es entrada
        inválida (exit 2), nunca se ignora (HR-3)."""
        excl = [dict(EXCL[0], requires=["gpu"])]
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp), build_policy(tmp, excl, tmp))
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("fail-closed", proc.stderr)
            self.assertIn("gpu", proc.stderr)

    def test_capability_declaration_does_not_waive_the_reason(self):
        """Declarar `requires` no exime de motivo: sigue siendo una diferencia
        declarada y exige `reason` (no relaja el gate)."""
        excl = [{"command": "python3 personal/ardelperal/ci-pattern/assets/x/check_x.py*",
                 "side": "ci", "requires": ["docker"]}]
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp), build_policy(tmp, excl, tmp))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("sin 'reason'", proc.stdout)

    def test_exclusion_without_reason_is_a_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            excl = EXCL + [{"command": "bash testing/suites/a/test-a.sh", "side": "any"}]
            proc = run_gate(build_repo(tmp), build_policy(tmp, excl, tmp))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("sin 'reason'", proc.stdout)

    def test_exclusion_matching_nothing_is_a_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            excl = EXCL + [{"command": "ghost-cmd", "side": "ci", "reason": "x"}]
            proc = run_gate(build_repo(tmp), build_policy(tmp, excl, tmp))
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("no coincide con ningún comando", proc.stdout)

    def test_empty_job_fails_closed(self):
        """Un job sin comandos de gate es sujeto vacío: exit 2, nunca PASS."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp, workflow=WORKFLOW.format(
                job="unit", steps="      - name: eco\n        run: echo hi\n"))
            proc = run_gate(repo, build_policy(tmp, EXCL, repo))
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("sujeto vacío", proc.stderr)

    def test_missing_entrypoint_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp)
            (repo / "testing" / "run-unit-suites.sh").unlink()
            proc = run_gate(repo, build_policy(tmp, EXCL, repo))
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("no existe", proc.stderr)

    def test_job_must_appear_exactly_once(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = build_repo(tmp, workflow=WORKFLOW.format(job="no-existe", steps=CI_STEPS))
            proc = run_gate(repo, build_policy(tmp, EXCL, repo))
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("exactamente una vez", proc.stderr)

    def test_dogfooding_this_repository_passes(self):
        """Dogfooding #196: el punto de entrada de este repo es
        testing/run-unit-suites.sh y su política declara las exclusiones
        CI-only con reason (HR-4 enmendada según #156)."""
        proc = subprocess.run(
            [sys.executable, str(GATE), "--repo", str(REPO_ROOT), "--policy", str(REAL_POLICY)],
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("LOCAL/CI PARITY OK", proc.stdout)

    def test_real_policy_exclusions_all_have_reasons(self):
        policy = json.loads(REAL_POLICY.read_text(encoding="utf-8"))
        self.assertTrue(policy["exclusions"], "el dogfooding debe declarar exclusiones")
        for exc in policy["exclusions"]:
            self.assertTrue(exc.get("reason", "").strip(), exc)


    def test_shorthand_block_siblings_are_not_script_lines(self):
        """`- run: |` con claves hermanas del step (shell:, env:,
        working-directory: a la columna de `run`) no las absorbe como
        líneas del script: sin comandos CI fantasma (hallazgo del auditor)."""
        wf = ("jobs:\n"
              "  unit:\n"
              "  steps:\n"
              "      - run: |\n"
              "          bash cmd/a.sh\n"
              "        shell: bash\n"
              "        env:\n"
              "          FOO: bar\n"
              "        working-directory: ./x\n")
        ep = 'if bash cmd/a.sh; then\n  echo PASS\nfi\n'
        with tempfile.TemporaryDirectory() as tmp:
            proc = run_gate(build_repo(tmp, workflow=wf, entrypoint=ep),
                            build_policy(tmp, [], tmp))
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertNotIn("shell", proc.stdout)
            self.assertNotIn("FOO", proc.stdout)



if __name__ == "__main__":
    unittest.main(verbosity=2)
