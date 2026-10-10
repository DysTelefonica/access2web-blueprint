#!/usr/bin/env python3
# ci-pattern asset tests — host readback drift check (DysTelefonica/team-skills#138)
"""Executable suite for the host readback drift check (black-box, like CI).
Drift cases exit 1 naming the rule; fail-closed cases exit non-zero (HR-3)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_host_drift.py"
EXAMPLE_CONTRACT = ASSET / "host-contract.example.json"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
KINDS = ("repo", "labels", "branch-protection", "rulesets")


def base_contract() -> dict:
    return {
        "contract_version": 1,
        "labels": [{"name": "status:approved", "class": "host-enforced"}],
        "required_checks": [{"name": "branch-name", "class": "host-enforced"}],
        "merge_methods": {
            "allow_merge_commit": {"declared": True, "class": "host-enforced"},
            "allow_squash_merge": {"declared": False, "class": "host-enforced"},
            "allow_rebase_merge": {"declared": False, "class": "host-enforced"},
        },
        "protection": {
            "enforce_admins": {"declared": True, "class": "host-enforced"},
            "required_conversation_resolution": {"declared": True, "class": "host-enforced"},
        },
        "rulesets": [{"name": "default-branch", "enforcement": "active", "class": "host-enforced"}],
    }


def base_snapshots(**overrides: object) -> dict[str, str]:
    snaps: dict[str, object] = {
        "repo": {"allow_merge_commit": True, "allow_squash_merge": False, "allow_rebase_merge": False},
        "labels": [{"name": "status:approved"}],
        "branch-protection": {
            "enforce_admins": {"enabled": True},
            "required_conversation_resolution": {"enabled": True},
            "required_status_checks": {"contexts": ["branch-name"],
                                       "checks": [{"context": "branch-name", "app_id": -1}]},
        },
        "rulesets": [{"name": "default-branch", "enforcement": "active"}],
    }
    snaps.update(overrides)
    return {kind: json.dumps(snaps[kind]) for kind in KINDS}


def run_gate(contract: dict | str, snapshots: dict[str, str]) -> subprocess.CompletedProcess:
    """Run the script as CI would; a missing kind omits its ``--snapshot`` flag."""
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        contract_path = root / "contract.json"
        contract_path.write_text(
            contract if isinstance(contract, str) else json.dumps(contract, indent=2), encoding="utf-8")
        argv = [sys.executable, str(SCRIPT), "--contract", str(contract_path)]
        for kind in KINDS:
            if kind not in snapshots:
                continue
            path = root / f"{kind}.json"
            path.write_text(snapshots[kind], encoding="utf-8")
            argv += ["--snapshot", f"{kind}={path}"]
        return subprocess.run(argv, capture_output=True, text=True)


class DriftTests(unittest.TestCase):
    def test_declared_label_absent_from_host_exits_one(self) -> None:
        proc = run_gate(base_contract(), base_snapshots(labels=[]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("label 'status:approved'", proc.stdout)
        self.assertIn("absent from the host", proc.stdout)

    def test_host_enforced_key_absent_from_snapshot_fails_closed(self) -> None:
        contract = base_contract()
        contract["protection"]["allow_force_pushes"] = {"declared": False, "class": "host-enforced"}
        bp = {"enforce_admins": {"enabled": True}}
        proc = run_gate(contract, base_snapshots(**{"branch-protection": bp}))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'allow_force_pushes' is declared host-enforced but absent", proc.stderr)
        self.assertIn("absence is doubt", proc.stderr)

    def test_host_merge_method_not_declared_exits_one(self) -> None:
        contract = base_contract()
        contract["merge_methods"] = {"allow_merge_commit": {"declared": True, "class": "host-enforced"}}
        repo = {"allow_merge_commit": True, "allow_squash_merge": True, "allow_rebase_merge": False}
        proc = run_gate(contract, base_snapshots(repo=repo))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("merge method 'allow_squash_merge' is enabled on the host but not declared", proc.stdout)

    def test_documented_only_rules_listed_without_failure(self) -> None:
        contract = base_contract()
        contract["labels"].append({"name": "status:blocked", "class": "documented-only"})
        contract["protection"]["required_conversation_resolution"] = {
            "declared": True, "class": "documented-only"}
        bp = json.loads(base_snapshots()["branch-protection"])
        bp["required_conversation_resolution"] = {"enabled": False}
        proc = run_gate(contract, base_snapshots(**{"branch-protection": bp}))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("documented-only  label 'status:blocked'", proc.stdout)
        self.assertIn("documented-only  protection flag 'required_conversation_resolution'", proc.stdout)
        self.assertIn("VERDICT: PASS", proc.stdout)


class FailClosedTests(unittest.TestCase):
    def test_missing_unreadable_snapshot_and_bad_contract_exit_two(self) -> None:
        missing = base_snapshots()
        del missing["branch-protection"]
        proc = run_gate(base_contract(), missing)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("missing snapshots: ['branch-protection']", proc.stderr)
        unreadable = base_snapshots()
        unreadable["repo"] = "{not json"  # verbatim: not valid JSON
        proc = run_gate(base_contract(), unreadable)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertIn("not valid JSON", proc.stderr)
        bad_class = base_contract()
        bad_class["labels"][0]["class"] = "maybe"
        proc = run_gate(bad_class, base_snapshots())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)


# 403 real de plan (DysTelefonica/cadete, 2026-10-06), capturado con una
# llamada de solo lectura; el gate lo registra como "unavailable", no error.
PLAN_403 = {
    "message": "Upgrade to GitHub Pro or make this repository public to enable this feature.",
    "documentation_url": "https://docs.github.com/rest/branches/branch-protection#get-branch-protection",
    "status": "403",
}


class ProtectedBranchesContractTests(unittest.TestCase):
    """#303: las ramas protegidas por contrato son dato validado (HR-34)."""

    def test_protected_branches_shape_is_validated(self) -> None:
        contract = base_contract()
        contract["protected_branches"] = ["main"]
        proc = run_gate(contract, base_snapshots())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        for bad in ([], ["main", 5], "main"):
            contract["protected_branches"] = bad
            proc = run_gate(contract, base_snapshots())
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("protected_branches", proc.stderr)


class HostCapabilitiesTests(unittest.TestCase):
    """#303: la capacidad del host es un dato declarado en el contrato
    (host_capabilities) y verificado por relectura; el 403 de plan de un
    repo privado se registra como `unavailable`, no como error."""

    def private_contract(self) -> dict:
        # Repo privado sin plan: solo lo que el plan gratuito sí aplica
        # (labels y merge methods) queda host-enforced.
        return {
            "contract_version": 1,
            "host_capabilities": {
                "branch_protection": "unavailable",
                "rulesets": "unavailable",
                "required_checks": "unavailable",
            },
            "labels": [{"name": "status:approved", "class": "host-enforced"}],
            "required_checks": [{"name": "branch-name", "class": "documented-only"}],
            "merge_methods": {"allow_merge_commit": {"declared": True, "class": "host-enforced"}},
            "protection": {"enforce_admins": {"declared": True, "class": "documented-only"}},
            "rulesets": [],
        }

    def private_snapshots(self) -> dict[str, str]:
        snaps = base_snapshots()
        snaps["branch-protection"] = json.dumps(PLAN_403)
        snaps["rulesets"] = json.dumps(PLAN_403)
        return snaps

    def test_plan_403_recorded_as_unavailable_passes(self) -> None:
        # El 403 de plan no es error: es la relectura que registra la
        # capacidad como unavailable; el resto del host sigue verificado.
        proc = run_gate(self.private_contract(), self.private_snapshots())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("unavailable", proc.stdout)

    def test_plan_403_with_declared_available_is_drift(self) -> None:
        # Sin host_capabilities (o declarando available), el 403 de plan es
        # drift que nombra la capacidad — no un error opaco de sintaxis.
        proc = run_gate(base_contract(), self.private_snapshots())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Upgrade to GitHub Pro", proc.stdout + proc.stderr)
        self.assertIn("available", proc.stdout + proc.stderr)

    def test_declared_unavailable_but_payload_available_is_drift(self) -> None:
        # Dirección contraria: el contrato dice unavailable y la relectura
        # muestra la capacidad real (bidireccional).
        contract = self.private_contract()
        contract["protection"] = {"enforce_admins": {"declared": True, "class": "documented-only"}}
        proc = run_gate(contract, base_snapshots())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("branch_protection", proc.stdout + proc.stderr)
        self.assertIn("unavailable", proc.stdout + proc.stderr)

    def test_host_enforced_rule_in_unavailable_area_is_drift(self) -> None:
        # Una regla host-enforced en un área que el host no puede aplicar
        # es contradicción del contrato: drift nombrando la regla.
        contract = self.private_contract()
        contract["protection"] = {"enforce_admins": {"declared": True, "class": "host-enforced"}}
        proc = run_gate(contract, self.private_snapshots())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("enforce_admins", proc.stdout + proc.stderr)

    def test_capability_shape_is_fail_closed(self) -> None:
        # Valor fuera del enum y clave desconocida son duda (exit 2).
        for caps in ({"branch_protection": "maybe"},
                     {"branch-protection": "available"},
                     {"branch_protection": None}):
            contract = self.private_contract()
            contract["host_capabilities"] = caps
            proc = run_gate(contract, self.private_snapshots())
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("fail-closed", proc.stderr)


class RunnerEnforcedTests(unittest.TestCase):
    """#303: tercera clase de HR-34. El control lo aplica un workflow en un
    runner propio declarado en el contrato; solo es válida en un área cuya
    capacidad el host declara `unavailable`, y exige `runner.label`."""

    def runner_contract(self) -> dict:
        contract = HostCapabilitiesTests().private_contract()
        contract["host_capabilities"]["branch_protection"] = "unavailable"
        contract["protection"] = {"enforce_admins": {"declared": False, "class": "runner-enforced"}}
        contract["runner"] = {"label": "cadete-oracle-arm64"}
        return contract

    def plan403_snapshots(self) -> dict[str, str]:
        return HostCapabilitiesTests().private_snapshots()

    def test_runner_enforced_in_unavailable_area_with_label_passes(self) -> None:
        proc = run_gate(self.runner_contract(), self.plan403_snapshots())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("runner-enforced", proc.stdout)
        self.assertIn("cadete-oracle-arm64", proc.stdout)

    def test_runner_enforced_with_available_capability_is_drift(self) -> None:
        # El host sí puede aplicar el área (sin límite de plan): la regla
        # debe ser host-enforced, no runner-enforced. Drift con nombre.
        contract = self.runner_contract()
        contract["host_capabilities"] = {"branch_protection": "available",
                                         "rulesets": "unavailable",
                                         "required_checks": "unavailable"}
        proc = run_gate(contract, base_snapshots())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("enforce_admins", proc.stdout + proc.stderr)
        self.assertIn("runner-enforced", proc.stdout + proc.stderr)
        self.assertIn("available", proc.stdout + proc.stderr)

    def test_runner_enforced_without_runner_label_fails_closed(self) -> None:
        contract = self.runner_contract()
        del contract["runner"]
        proc = run_gate(contract, self.plan403_snapshots())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("runner", proc.stderr)

    def test_runner_label_shape_is_fail_closed(self) -> None:
        for runner in ({"label": ""}, {"label": 5}, {"label": "cadete", "extra": 1}):
            contract = self.runner_contract()
            contract["runner"] = runner
            proc = run_gate(contract, self.plan403_snapshots())
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("fail-closed", proc.stderr)

    def test_runner_enforced_label_rule_is_drift(self) -> None:
        # Las etiquetas las aplica el host en cualquier plan: una etiqueta
        # runner-enforced es contradicción, no una clase válida ahí.
        contract = self.runner_contract()
        contract["labels"] = [{"name": "status:approved", "class": "runner-enforced"}]
        proc = run_gate(contract, self.plan403_snapshots())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("status:approved", proc.stdout + proc.stderr)


class VacuousAndAdversarialTests(unittest.TestCase):
    """#185: sujeto vacío, cuerpos de error de la API y reglas no declaradas.

    HR-18: una medición degenerada no es un PASS; HR-3: una lectura
    fallida del host es duda, no evidencia."""

    def test_empty_contract_with_empty_host_fails_closed(self) -> None:
        # Un contrato sin ninguna entrada aplicada por el host y snapshots
        # sin nada que verificar es una medición vacía: nunca PASS.
        proc = run_gate({"contract_version": 1},
                        base_snapshots(repo={}, labels=[], **{"branch-protection": {}, "rulesets": []}))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertIn("vacuous", proc.stderr)

    def test_api_error_snapshot_fails_closed_naming_file(self) -> None:
        # Un snapshot con forma de error de la API ('message' sin los
        # campos esperados) falla cerrado y nombra el fichero.
        err = {"message": "Not Found", "documentation_url": "https://docs.github.com/rest"}
        proc = run_gate(base_contract(), base_snapshots(**{"branch-protection": err}))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertIn("error", proc.stderr.lower())
        self.assertIn("branch-protection", proc.stderr)
        self.assertIn(".json", proc.stderr)

    def test_undeclared_active_ruleset_is_a_finding(self) -> None:
        # Un ruleset activo en el host que el contrato no declara como
        # host-enforced es drift (salida en ambas direcciones).
        rs = [{"name": "default-branch", "enforcement": "active"},
              {"name": "extra-ruleset", "enforcement": "active"}]
        proc = run_gate(base_contract(), base_snapshots(rulesets=rs))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("ruleset 'extra-ruleset'", proc.stdout)
        self.assertIn("active on the host but not declared", proc.stdout)


class EvaluateRulesetTests(unittest.TestCase):
    """#234: un ruleset del host no declarado es drift en cualquier modo
    distinto de 'disabled' (evaluate registra evaluaciones y suele preceder
    a la activacion), y un valor de 'enforcement' desconocido es duda que
    falla cerrado (HR-3)."""

    def test_undeclared_evaluate_ruleset_is_a_finding(self) -> None:
        rs = [{"name": "default-branch", "enforcement": "active"},
              {"name": "dry-run", "enforcement": "evaluate"}]
        proc = run_gate(base_contract(), base_snapshots(rulesets=rs))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("ruleset 'dry-run'", proc.stdout)
        self.assertIn("evaluate", proc.stdout)
        self.assertIn("not declared", proc.stdout)

    def test_undeclared_disabled_ruleset_is_not_a_finding(self) -> None:
        rs = [{"name": "default-branch", "enforcement": "active"},
              {"name": "shadow", "enforcement": "disabled"}]
        proc = run_gate(base_contract(), base_snapshots(rulesets=rs))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)

    def test_undeclared_ruleset_without_enforcement_is_a_finding(self) -> None:
        # Clave 'enforcement' ausente: el host muestra un ruleset que este
        # control no puede clasificar; no es disabled, luego es hallazgo.
        rs = [{"name": "default-branch", "enforcement": "active"},
              {"name": "mystery"}]
        proc = run_gate(base_contract(), base_snapshots(rulesets=rs))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("ruleset 'mystery'", proc.stdout)

    def test_unknown_enforcement_value_fails_closed_naming_value(self) -> None:
        # Un valor desconocido ('bogus') y un null explicito son duda: el
        # host aplica algo que este control no sabe interpretar (HR-3).
        for value in ("bogus", None):
            with self.subTest(enforcement=value):
                rs = [{"name": "default-branch", "enforcement": "active"},
                      {"name": "weird", "enforcement": value}]
                proc = run_gate(base_contract(), base_snapshots(rulesets=rs))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("fail-closed", proc.stderr)
                self.assertIn("weird", proc.stderr)
                self.assertIn("rulesets", proc.stderr)
        proc = run_gate(base_contract(), base_snapshots(
            rulesets=[{"name": "default-branch", "enforcement": "active"},
                      {"name": "weird", "enforcement": "bogus"}]))
        self.assertIn("bogus", proc.stderr)

    def test_contract_may_declare_evaluate_ruleset(self) -> None:
        contract = base_contract()
        contract["rulesets"].append({"name": "dry-run", "enforcement": "evaluate",
                                     "class": "host-enforced"})
        rs = [{"name": "default-branch", "enforcement": "active"},
              {"name": "dry-run", "enforcement": "evaluate"}]
        proc = run_gate(contract, base_snapshots(rulesets=rs))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("ruleset 'dry-run' matches the host", proc.stdout)


class ExampleContractTests(unittest.TestCase):
    def test_example_contract_passes_against_matching_host(self) -> None:
        argv = [sys.executable, str(SCRIPT), "--contract", str(EXAMPLE_CONTRACT)]
        for kind in KINDS:
            argv += ["--snapshot", f"{kind}={FIXTURES / (kind + '.json')}"]
        proc = subprocess.run(argv, capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("documented-only  label 'status:blocked'", proc.stdout)
        self.assertIn("documented-only  protection flag 'required_conversation_resolution'", proc.stdout)
        self.assertIn("VERDICT: PASS", proc.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
