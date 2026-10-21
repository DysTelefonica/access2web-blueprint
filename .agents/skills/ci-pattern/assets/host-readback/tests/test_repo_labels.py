#!/usr/bin/env python3
"""#237: check_host_drift detecta la ausencia de chain:partial en el host.

Usa el contrato del propio repo (.github/host-contract.json) y fixtures
de snapshots de etiquetas. El snapshot sin chain:partial produce un
hallazgo de drift; con ella, pasa. RED: el contrato no existe aún.
"""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[6]
DRIFT = REPO_ROOT / "personal" / "ardelperal" / "ci-pattern" / "assets" / "host-readback" / "check_host_drift.py"
CONTRACT = REPO_ROOT / ".github" / "host-contract.json"

SKILL = "personal/ardelperal/ci-pattern"


def _branch_protection_snapshot(contract):
    """Snapshot sintetico de branch-protection construido desde los
    ``required_checks`` declarados como ``host-enforced`` en el contrato.

    El test mide coherencia interna entre contrato y snapshot sintetico,
    no la realidad del host (eso es trabajo de ``sync_governance_labels.sh``
    y del ``check_governed_labels.sh`` en tiempo de merge, fuera del unit
    suite). Cuando el contrato declara un ``required_check`` como
    ``host-enforced``, este snapshot lo refleja para que el gate no
    reporte drift por checks ya declarados (regla del #900, HR-34):
    un ``required_check`` declarado en el contrato y ausente del snapshot
    es exactamente la clase de falso positivo que este test evita.
    """
    contexts = [item["name"] for item in contract.get("required_checks", [])
                if isinstance(item, dict) and item.get("class") == "host-enforced"]
    return {
        "required_status_checks": {
            "contexts": contexts,
            "checks": [{"context": c} for c in contexts],
        },
    }


def _run_drift(labels_snapshot, branch_protection=None):
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    if branch_protection is None:
        branch_protection = _branch_protection_snapshot(contract)
    with tempfile.TemporaryDirectory() as tmp:
        snap_dir = Path(tmp)
        (snap_dir / "labels.json").write_text(json.dumps(labels_snapshot), encoding="utf-8")
        (snap_dir / "repo.json").write_text("{}", encoding="utf-8")
        (snap_dir / "branch-protection.json").write_text(
            json.dumps(branch_protection), encoding="utf-8")
        (snap_dir / "rulesets.json").write_text("[]", encoding="utf-8")
        return subprocess.run(
            [sys.executable, str(DRIFT),
             "--contract", str(CONTRACT),
             "--snapshot", f"repo={snap_dir / 'repo.json'}",
             "--snapshot", f"labels={snap_dir / 'labels.json'}",
             "--snapshot", f"branch-protection={snap_dir / 'branch-protection.json'}",
             "--snapshot", f"rulesets={snap_dir / 'rulesets.json'}"],
            capture_output=True, text=True, check=False,
        )


class RepoHostContractTests(unittest.TestCase):

    def test_contract_exists_and_declares_chain_partial(self):
        """El contrato existe y declara chain:partial como host-enforced."""
        self.assertTrue(CONTRACT.is_file(), ".github/host-contract.json no existe")
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        labels = {item["name"]: item["class"] for item in contract.get("labels", [])}
        self.assertEqual(labels.get("chain:partial"), "host-enforced")
        self.assertEqual(labels.get("status:approved"), "host-enforced")

    @staticmethod
    def _contract_labels():
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        return [item["name"] for item in contract.get("labels", []) if item.get("class") == "host-enforced"]

    def test_drift_without_chain_partial_produces_finding(self):
        """Snapshot sin chain:partial → drift (hallazgo HR-34)."""
        self.assertTrue(CONTRACT.is_file(), "ejecutar después de crear el contrato")
        labels = [n for n in self._contract_labels() if n != "chain:partial"]
        snapshot = [{"name": n} for n in labels]
        proc = _run_drift(snapshot)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("chain:partial", proc.stdout)

    def test_drift_with_chain_partial_passes(self):
        """Snapshot con chain:partial → sin drift para labels."""
        self.assertTrue(CONTRACT.is_file(), "ejecutar después de crear el contrato")
        labels = self._contract_labels()
        snapshot = [{"name": n} for n in labels]
        proc = _run_drift(snapshot)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_label_descriptions_fit_github_limit(self):
        """Todas las descripciones del contrato caben en el limite de la API
        de GitHub (100 caracteres) (#900 follow-up).

        ``sync-governance-labels.sh`` llama al endpoint de edicion de labels
        de GitHub con la ``description`` que el contrato declara. GitHub
        rechaza la peticion con
        ``"description is too long (maximum is 100 characters)"`` y deja la
        foto del repo sin la etiqueta, divergiendo del contrato. Caso real
        (octubre 2026): ``size:exception`` declarada con 110 caracteres y la
        etiqueta real que el mantenedor habia creado llevaba 81. El test
        recorre TODAS las descripciones declaradas (no solo las
        ``host-enforced``: una etiqueta ``documented-only`` que sincronice la
        foto del repo en algun punto cae en el mismo error) y falla con
        nombre, largo y texto si alguna supera el limite.
        """
        self.assertTrue(CONTRACT.is_file(), "ejecutar despues de crear el contrato")
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        labels = contract.get("labels", [])
        self.assertTrue(labels, "el contrato debe declarar al menos una etiqueta")
        GITHUB_LABEL_DESCRIPTION_LIMIT = 100
        offenders = []
        for entry in labels:
            if not isinstance(entry, dict):
                continue
            name = entry.get("name", "?")
            description = entry.get("description", "")
            if not isinstance(description, str):
                offenders.append(
                    f"{name!r}: descripcion no es string (es {type(description).__name__})"
                )
                continue
            n = len(description)
            if n > GITHUB_LABEL_DESCRIPTION_LIMIT:
                offenders.append(
                    f"{name!r}: {n} caracteres (limite {GITHUB_LABEL_DESCRIPTION_LIMIT}): {description!r}"
                )
        self.assertEqual(offenders, [], "\n".join(offenders) if offenders else "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
