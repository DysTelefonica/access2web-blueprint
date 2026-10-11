#!/usr/bin/env python3
"""Suite del meta-gate HR→gate (#191).

El meta-gate (`../check_hr_matrix.py`) se ejecuta como proceso contra un
árbol de skill falso construido en tmp: cubre el caso positivo, los
negativos (HR sin registro, HR inexistente, gate con fichero/test
ausente, manual sin reason/issue) y los degenerados (matriz ausente o
vacía, SKILL.md ausente). Ejecución: `python3 test_hr_matrix.py` o el
shim `run.sh` (que además hace dogfooding contra la skill real).
"""

import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_hr_matrix.py"


def build_skill(tmp, hrs, matrix, matrix_name="hr-gate-matrix.json"):
    skill = Path(tmp) / "ci-pattern"
    (skill / "references").mkdir(parents=True)
    (skill / "assets" / "gate").mkdir(parents=True)
    (skill / "assets" / "gate" / "gate.py").write_text("# gate\n", encoding="utf-8")
    (skill / "assets" / "gate" / "test_gate.py").write_text("# test\n", encoding="utf-8")
    hr_lines = "\n".join(f"- **{h} — regla de prueba {h}**: texto." for h in hrs)
    (skill / "SKILL.md").write_text(
        "# Skill de prueba\n\n" + hr_lines + "\n", encoding="utf-8"
    )
    if matrix is not None:
        (skill / "references" / matrix_name).write_text(
            json.dumps(matrix, indent=2) + "\n", encoding="utf-8"
        )
    return skill


def run_gate(skill):
    return subprocess.run(
        [sys.executable, str(GATE), str(skill)],
        capture_output=True, text=True, check=False,
    )


def entry(h, enforcement, **kw):
    base = {"id": h, "enforcement": enforcement}
    base.update(kw)
    return base


class HrGateMatrixTests(unittest.TestCase):
    def test_positive_full_matrix(self):
        """Positivo: toda HR registrada, gates con ficheros reales."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/test_gate.py"),
                entry("HR-2", "manual", reason="comportamiento del agente", issue=191),
            ]}
            skill = build_skill(tmp, ["HR-1", "HR-2"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("gate 1", proc.stdout)
            self.assertIn("manual 1", proc.stdout)

    def test_hr_in_skill_missing_from_matrix(self):
        """Negativo: una HR de SKILL.md sin entrada en la matriz."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/test_gate.py"),
            ]}
            skill = build_skill(tmp, ["HR-1", "HR-2"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("HR-2", proc.stdout)

    def test_matrix_cites_unknown_hr(self):
        """Negativo: la matriz cita una HR que no está en SKILL.md."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/test_gate.py"),
                entry("HR-99", "manual", reason="x", issue=191),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("HR-99", proc.stdout)

    def test_gate_entry_with_missing_asset(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/falta.py", test="assets/gate/test_gate.py"),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("falta.py", proc.stdout)

    def test_gate_entry_with_missing_test(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/falta_test.py"),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("falta_test.py", proc.stdout)

    def test_manual_without_reason_and_issue(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [entry("HR-1", "manual")]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("reason", proc.stdout)
            self.assertIn("issue", proc.stdout)

    def test_empty_rules_with_hrs_declared(self):
        """Degenerado: matriz sin reglas cuando SKILL.md declara HR."""
        with tempfile.TemporaryDirectory() as tmp:
            skill = build_skill(tmp, ["HR-1"], {"rules": []})
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)

    def test_missing_matrix_file(self):
        """Degenerado: sin fichero de matriz → exit 2, nunca verde."""
        with tempfile.TemporaryDirectory() as tmp:
            skill = build_skill(tmp, ["HR-1"], None)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 2)

    def test_missing_skill_md(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = Path(tmp) / "vacio"
            skill.mkdir()
            (skill / "references").mkdir()
            (skill / "references" / "hr-gate-matrix.json").write_text(
                json.dumps({"rules": []}), encoding="utf-8"
            )
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 2)

    def test_process_with_reason_passes(self):
        """Nueva clase process: reason exigido, issue NO exigido."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "process", reason="Regla de diseño de gates."),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("process 1", proc.stdout)

    def test_process_without_reason_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [entry("HR-1", "process")]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("reason", proc.stdout)

    def test_process_with_issue_still_passes(self):
        """Si una process lleva issue (legado), no es violación."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "process", reason="Comportamiento del agente.", issue=191),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_manual_issue_closed_fails_with_open_issues_file(self):
        """Con --open-issues-file, una manual cuya issue no está abierta es violación."""
        with tempfile.TemporaryDirectory() as tmp:
            issues_file = Path(tmp) / "open-issues.txt"
            issues_file.write_text("200\n201\n", encoding="utf-8")
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/test_gate.py"),
                entry("HR-2", "manual", reason="Pendiente el gate de contrato.", issue=193),
            ]}
            skill = build_skill(tmp, ["HR-1", "HR-2"], matrix)
            proc = subprocess.run(
                [sys.executable, str(GATE), str(skill),
                 "--open-issues-file", str(issues_file)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 1)
            self.assertIn("cerrada", proc.stdout)
            self.assertIn("193", proc.stdout)

    def test_manual_issue_open_passes_with_open_issues_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            issues_file = Path(tmp) / "open-issues.txt"
            issues_file.write_text("193\n", encoding="utf-8")
            matrix = {"rules": [
                entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/test_gate.py"),
                entry("HR-2", "manual", reason="Pendiente el gate de contrato.", issue=193),
            ]}
            skill = build_skill(tmp, ["HR-1", "HR-2"], matrix)
            proc = subprocess.run(
                [sys.executable, str(GATE), str(skill),
                 "--open-issues-file", str(issues_file)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_process_not_affected_by_open_issues_check(self):
        """Las process no se validan contra la lista de issues abiertas."""
        with tempfile.TemporaryDirectory() as tmp:
            issues_file = Path(tmp) / "open-issues.txt"
            issues_file.write_text("193\n", encoding="utf-8")
            matrix = {"rules": [
                entry("HR-1", "process", reason="Comportamiento del agente."),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = subprocess.run(
                [sys.executable, str(GATE), str(skill),
                 "--open-issues-file", str(issues_file)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_without_open_issues_file_report_declares_check_not_performed(self):
        """#235: sin lista de issues abiertas, el informe lo declara
        explicitamente; nunca sugiere que la comprobación se hizo."""
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [
                entry("HR-1", "manual", reason="Pendiente el gate.", issue=193),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("comprobación de issues no realizada", proc.stdout)
            self.assertNotIn("issue abierta", proc.stdout)

    def test_with_open_issues_file_report_declares_check_performed(self):
        """#235: con lista, el informe declara la comprobación hecha."""
        with tempfile.TemporaryDirectory() as tmp:
            issues_file = Path(tmp) / "open-issues.txt"
            issues_file.write_text("193\n", encoding="utf-8")
            matrix = {"rules": [
                entry("HR-1", "manual", reason="Pendiente el gate.", issue=193),
            ]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = subprocess.run(
                [sys.executable, str(GATE), str(skill),
                 "--open-issues-file", str(issues_file)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("issues abiertas verificadas", proc.stdout)
            self.assertNotIn("no realizada", proc.stdout)

    def test_real_skill_dogfooding(self):
        """Dogfooding: la skill real pasa su propio meta-gate y los
        recuentes son consistentes con SKILL.md (sin snapshots: cambiar
        el número de HRs no rompe el test)."""
        skill = Path(__file__).resolve().parents[3]
        cmd = [sys.executable, str(GATE), str(skill)]
        # #235: en CI, un paso previo genera la lista de issues abiertas y
        # la pasa por env; el dogfooding la inyecta al gate si existe.
        env_file = os.environ.get("HR_MATRIX_OPEN_ISSUES_FILE")
        if env_file and Path(env_file).is_file():
            cmd += ["--open-issues-file", env_file]
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        if env_file and Path(env_file).is_file():
            self.assertNotIn("comprobación de issues no realizada", proc.stdout)
        else:
            self.assertIn("comprobación de issues no realizada", proc.stdout)
        m = re.search(r"gate (\d+), manual (\d+), process (\d+)", proc.stdout)
        self.assertIsNotNone(m, proc.stdout)
        gate_n, manual_n, process_n = int(m.group(1)), int(m.group(2)), int(m.group(3))
        hr_ids = set(re.findall(
            r"^- \*\*(HR-\d+)\b",
            (skill / "SKILL.md").read_text(encoding="utf-8"),
            re.MULTILINE,
        ))
        self.assertEqual(gate_n + manual_n + process_n, len(hr_ids))
        # Los tres gates entregados (HR-29/34/35) no se caen en silencio.
        self.assertGreaterEqual(gate_n, 3)

    def test_real_skill_dogfooding_with_open_issues_env(self):
        """#235 modo CI: lista de issues abiertas pasada por env (aquí se
        genera offline desde la propia matriz; en CI la genera el paso
        previo con `gh issue list --state open`). El gate la consume y el
        informe declara la comprobación hecha."""
        skill = Path(__file__).resolve().parents[3]
        matrix = json.loads(
            (skill / "references" / "hr-gate-matrix.json").read_text(encoding="utf-8")
        )
        numbers = sorted({r["issue"] for r in matrix["rules"]
                          if r.get("enforcement") == "manual" and isinstance(r.get("issue"), int)})
        # La matriz puede no tener ya entradas manual con issue (olas de
        # mecanización): el perro sigue echando a andar con la lista vacía.
        with tempfile.TemporaryDirectory() as tmp:
            issues_file = Path(tmp) / "open-issues.txt"
            issues_file.write_text("".join(f"{n}\n" for n in numbers), encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(GATE), str(skill),
                 "--open-issues-file", str(issues_file)],
                capture_output=True, text=True, check=False,
            )
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("issues abiertas verificadas", proc.stdout)
        self.assertNotIn("comprobación de issues no realizada", proc.stdout)

    def test_ci_step_limit_1000_and_fails_on_empty_list(self):
        """#235 corrección: el paso previo pide --limit 1000 (el default de
        `gh issue list` es 30; al superarlas faltarían issues antiguas y el
        gate daría falsos fallos) y falla si la lista sale vacía, porque un
        resultado vacío es un fallo de la API, no cero issues abiertas."""
        wf = Path(__file__).resolve().parents[6] / ".github" / "workflows" / "tests.yml"
        text = wf.read_text(encoding="utf-8")
        step = text[text.index("Lista de issues abiertas"):text.index("Suites de smoke y unit")]
        self.assertIn("--limit 1000", step)
        self.assertIn('[ ! -s "$RUNNER_TEMP/open-issues.txt" ]', step)
        self.assertIn("HR-3", step)  # el fallo por lista vacía es fail-loud

    def test_gate_with_empty_asset_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [entry("HR-1", "gate", asset="", test="assets/gate/test_gate.py")]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("vacío", proc.stdout)

    def test_gate_with_empty_test_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            matrix = {"rules": [entry("HR-1", "gate", asset="assets/gate/gate.py", test="")]}
            skill = build_skill(tmp, ["HR-1"], matrix)
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("vacío", proc.stdout)

    def test_failing_test_command_reported(self):
        """Un gate cuyo test declarado falla al ejecutarse es violación."""
        with tempfile.TemporaryDirectory() as tmp:
            skill = build_skill(tmp, ["HR-1"], None)
            failing = skill / "assets" / "gate" / "failing_test.sh"
            failing.write_text("#!/usr/bin/env bash\nexit 3\n", encoding="utf-8")
            failing.chmod(0o755)
            matrix = {"rules": [entry("HR-1", "gate", asset="assets/gate/gate.py", test="assets/gate/failing_test.sh")]}
            (skill / "references" / "hr-gate-matrix.json").write_text(
                json.dumps(matrix, indent=2) + "\n", encoding="utf-8"
            )
            proc = run_gate(skill)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("falla", proc.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
