#!/usr/bin/env python3
"""#237 (correccion del auditor): sync-governance-labels.sh create-only por defecto.

El script emite comandos gh, nunca los ejecuta. Por defecto solo emite
``gh label create`` para las etiquetas del contrato que FALTAN en el host;
``gh label edit`` (que sobreescribiria color y descripcion de etiquetas
existentes) solo se emite con el flag explicito ``--update``. La lista
actual de etiquetas se toma de un fichero (salida de
``gh label list --json name``) o de una consulta en solo lectura.
La descripcion y el color de cada etiqueta salen del contrato.
"""

import json
import subprocess
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[6]
SCRIPT = REPO_ROOT / "scripts" / "sync-governance-labels.sh"

CHAIN_DESC = "Slice intermedio de una cadena apilada contra main: lleva Refs #N, nunca Closes"

CONTRACT = {
    "contract_version": 1,
    "labels": [
        {"name": "status:approved", "class": "host-enforced",
         "description": "Aprobada por el mantenedor para su implementacion", "color": "0e8a16"},
        {"name": "chain:partial", "class": "host-enforced",
         "description": CHAIN_DESC, "color": "1d76db"},
    ],
    "required_checks": [],
    "merge_methods": {},
    "protection": {},
    "rulesets": [],
}

# Forma de la salida de `gh label list --json name`: solo status:approved existe.
HOST_LABELS = [{"name": "status:approved"}]


def run_script(args, contract=CONTRACT, host_labels=HOST_LABELS, with_gh_stub=False):
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        contract_path = root / "contract.json"
        contract_path.write_text(json.dumps(contract), encoding="utf-8")
        labels_path = root / "labels.json"
        labels_path.write_text(json.dumps(host_labels), encoding="utf-8")
        log_path = root / "gh.log"
        env_gh = ""
        if with_gh_stub:
            # gh simulado: registra cada invocacion; `label list` devuelve el
            # snapshot; cualquier subcomando mutador falla y queda registrado.
            gh = root / "gh"
            gh.write_text(
                "#!/usr/bin/env bash\n"
                f'echo "$@" >> {json.dumps(str(log_path))}\n'
                'if [ "$1" = "label" ] && [ "$2" = "list" ]; then\n'
                f'  cat {json.dumps(str(labels_path))}\n'
                '  exit 0\n'
                'fi\n'
                'exit 1\n',
                encoding="utf-8")
            gh.chmod(0o755)
            env_gh = str(root)
        proc = subprocess.run(
            ["bash", str(SCRIPT), *args, "--contract", str(contract_path),
             "--labels-file", str(labels_path)],
            capture_output=True, text=True, check=False,
            env=None if not env_gh else {"PATH": f"{env_gh}:/usr/bin:/bin", "HOME": str(root)},
        )
        log = log_path.read_text(encoding="utf-8") if log_path.exists() else ""
        return proc, log


class CreateOnlyByDefaultTests(unittest.TestCase):

    def test_emits_create_only_for_missing_labels(self):
        """Sin --update: create solo para chain:partial; ningun edit."""
        proc, _ = run_script([])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn('gh label create "chain:partial"', proc.stdout)
        self.assertNotIn('gh label create "status:approved"', proc.stdout)
        self.assertNotIn("gh label edit", proc.stdout)

    def test_create_uses_contract_description_and_color(self):
        """La descripcion y el color del create salen del contrato."""
        proc, _ = run_script([])
        self.assertIn(CHAIN_DESC, proc.stdout)
        self.assertIn("--color \"1d76db\"", proc.stdout)

    def test_no_label_edit_anywhere_without_update_flag(self):
        """HR-7 de esta correccion: '|| gh label edit' jamas sin --update."""
        proc, _ = run_script([])
        self.assertNotIn("label edit", proc.stdout)


class UpdateFlagTests(unittest.TestCase):

    def test_update_emits_edit_for_existing_labels_only(self):
        """Con --update: edit para las existentes, create para las que faltan."""
        proc, _ = run_script(["--update"])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn('gh label edit "status:approved"', proc.stdout)
        self.assertIn('gh label create "chain:partial"', proc.stdout)
        self.assertIn(CHAIN_DESC, proc.stdout)


class ReadOnlySourceTests(unittest.TestCase):

    def test_current_labels_from_gh_stub_list_only(self):
        """Sin --labels-file, consulta gh label list (solo lectura); nunca muta."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(CONTRACT), encoding="utf-8")
            labels_path = root / "labels.json"
            labels_path.write_text(json.dumps(HOST_LABELS), encoding="utf-8")
            log_path = root / "gh.log"
            gh = root / "gh"
            gh.write_text(
                "#!/usr/bin/env bash\n"
                f'echo "$@" >> {json.dumps(str(log_path))}\n'
                'if [ "$1" = "label" ] && [ "$2" = "list" ]; then\n'
                f'  cat {json.dumps(str(labels_path))}\n'
                '  exit 0\n'
                'fi\n'
                'exit 1\n',
                encoding="utf-8")
            gh.chmod(0o755)
            proc = subprocess.run(
                ["bash", str(SCRIPT), "--contract", str(contract_path)],
                capture_output=True, text=True, check=False,
                env={"PATH": f"{root}:/usr/bin:/bin", "HOME": str(root)},
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn('gh label create "chain:partial"', proc.stdout)
            log = log_path.read_text(encoding="utf-8") if log_path.exists() else ""
            self.assertIn("label list", log)
            self.assertNotIn("label create", log)
            self.assertNotIn("label edit", log)


if __name__ == "__main__":
    unittest.main(verbosity=2)
