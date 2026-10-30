#!/usr/bin/env python3
"""El contrato del host del canon no promete lo que el host no aplica (R6).

Usa el contrato del propio repo (.github/host-contract.json) y fixtures
de snapshots del host. El snapshot sin chain:partial produce un
hallazgo de drift; con ella, pasa. RED: el contrato no existe aún.

R6 (#1093) anade el reverso: el canon es privado y sin plan Pro, asi que
la API de proteccion de rama y la de rulesets responden 403 de plan. El
contrato declara esa capacidad como `unavailable` y ninguna regla de esas
areas se declara `host-enforced`: el host no puede exigirlas. Las
instantaneas congeladas de abajo son la foto del host del 2026-10-23
(GETs de solo lectura); el veredicto del gate contra ellas es la prueba
de que el contrato no promete mas de lo que el host aplica.
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

# Cuerpo 403 de limite de plan que devuelve el host a un repo privado sin
# Pro en la proteccion de rama y en los rulesets (team-skills#330). Es
# EVIDENCIA de que el area no existe, no un error de lectura.
PLAN_403 = {
    "message": "Upgrade to GitHub Pro or make this repository public to enable this feature.",
    "status": "403",
}

# Ajustes del repo congelados el 2026-10-23 (GET /repos/DysTelefonica/team-skills).
REAL_REPO_SETTINGS = {
    "allow_merge_commit": True,
    "allow_squash_merge": True,
    "allow_rebase_merge": True,
    "allow_update_branch": False,
    "allow_auto_merge": False,
}

# Capacidad por area que gobierna que clase es admisible en esa area.
CAPABILITY_OF_AREA = {
    "required_checks": "required_checks",
    "protection": "branch_protection",
    "rulesets": "rulesets",
}


def _host_enforced_entries(contract, section):
    """Entradas `host-enforced` de una seccion del contrato."""
    if section == "protection":
        entries = list((contract.get("protection") or {}).values())
    else:
        entries = list(contract.get(section) or [])
    return [entry for entry in entries
            if isinstance(entry, dict) and entry.get("class") == "host-enforced"]


def _run_drift(labels_snapshot, branch_protection=None, repo=None, rulesets=None):
    if branch_protection is None:
        branch_protection = PLAN_403
    if repo is None:
        repo = REAL_REPO_SETTINGS
    if rulesets is None:
        rulesets = PLAN_403
    with tempfile.TemporaryDirectory() as tmp:
        snap_dir = Path(tmp)
        (snap_dir / "labels.json").write_text(json.dumps(labels_snapshot), encoding="utf-8")
        (snap_dir / "repo.json").write_text(json.dumps(repo), encoding="utf-8")
        (snap_dir / "branch-protection.json").write_text(
            json.dumps(branch_protection), encoding="utf-8")
        (snap_dir / "rulesets.json").write_text(json.dumps(rulesets), encoding="utf-8")
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

    def test_contract_declares_the_capability_of_every_area(self):
        """R6: la capacidad se declara en las tres areas. Sin la clave, el
        gate lee `available` por omision y el contrato promete por omision
        lo que el host no aplica (este repo: 403 de plan en proteccion de
        rama y en rulesets)."""
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        caps = contract.get("host_capabilities")
        self.assertIsInstance(caps, dict, "el contrato debe declarar host_capabilities")
        for area in ("branch_protection", "rulesets", "required_checks"):
            self.assertIn(area, caps, f"host_capabilities sin '{area}'")
            self.assertIn(caps[area], ("available", "unavailable"),
                          f"host_capabilities['{area}'] = {caps[area]!r} no es un valor del vocabulario")

    def test_contract_does_not_claim_host_enforcement_the_host_cannot_apply(self):
        """R6: en un area con capacidad `unavailable` ninguna regla se declara
        `host-enforced`. El host no exige checks, ni proteccion, ni rulesets
        aqui: prometerlo es una garantia que ningun host aplica."""
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        caps = contract.get("host_capabilities") or {}
        offenders = []
        for section, capability in CAPABILITY_OF_AREA.items():
            if caps.get(capability) != "unavailable":
                continue
            for entry in _host_enforced_entries(contract, section):
                offenders.append(f"{section}: {entry.get('name') or entry.get('declared')}")
        self.assertEqual(offenders, [],
                         "reglas declaradas host-enforced en un area que el host no aplica: "
                         + ", ".join(offenders))

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

    def test_contract_against_the_real_host_shape_does_not_drift(self):
        """R6: el contrato del canon, contrastado con la foto del host
        (ajustes del repo, etiquetas, 403 de plan en proteccion y rulesets),
        no deriva: el gate sale en verde."""
        snapshot = [{"name": n} for n in self._contract_labels()]
        proc = _run_drift(snapshot)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)
        self.assertNotIn("DRIFT", proc.stdout)

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
