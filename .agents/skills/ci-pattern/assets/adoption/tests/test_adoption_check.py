#!/usr/bin/env python3
"""Suite del gate `ci-pattern adoption check` (#276).

El gate es de solo lectura y stdlib: valida el contrato de adopción
`.github/ci-pattern-adoption.json` del consumer y comprueba cada fase del
porting-guide sobre el árbol real del repo destino.

Códigos de salida verificados: 0 fase cumple · 1 hallazgos nombrando cada
artefacto o condición que falta · 2 contrato ilegible, esquema inválido o
sujeto vacío. Ninguna fase pasa si la anterior no pasa.

Fixtures: repos falsos en directorios temporales (convención de
`assets/tests/test-ci-pattern-cli.py`); ninguno escribe fuera de `tmp`.
"""

import contextlib
import hashlib
import importlib.machinery
import importlib.util
import io
import json
import os
import re
import subprocess
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"
SKILL_DIR = SKILL_ROOT


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli",
        importlib.machinery.SourceFileLoader("ci_pattern_cli", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()


def run_cli(argv):
    out, err = io.StringIO(), io.StringIO()
    old = os.environ.copy()
    os.environ["CI_PATTERN_SKILL_DIR"] = str(SKILL_DIR)
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = CLI.main(argv)
    finally:
        os.environ.clear()
        os.environ.update(old)
    return code, out.getvalue(), err.getvalue()


def make_contract_repo(tmp, contract=None, contract_raw=None, complete_upto=0):
    """Repo falso con contrato de adopción y artefactos completos hasta la
    fase `complete_upto` (ambos inclusive). Devuelve (repo, contract_path).
    Los ficheros que las fases EJECUTAN (fase 2: árbol de tests limpio +
    allowlist; fase 4: workflow pineado y política de jobs) se crean
    siempre: son inertes para las fases 0-1 y los tests que ejercitan el
    fallo los sobreescriben o borran."""
    repo = Path(tmp) / "consumer"
    (repo / ".github").mkdir(parents=True)
    (repo / "docs" / "adoption").mkdir(parents=True)
    (repo / "tests").mkdir()
    (repo / "tests" / "test_smoke.sh").write_text("#!/bin/sh\necho ok\n", encoding="utf-8")
    (repo / ".github" / "env-isolation-allowlist.json").write_text(
        json.dumps({"env_vars": []}), encoding="utf-8")
    (repo / ".github" / "workflows").mkdir()
    (repo / ".github" / "workflows" / "ci.yml").write_text(
        "on:\n  pull_request:\njobs:\n  unit:\n    timeout-minutes: 5\n    concurrency: u\n    steps:\n"
        "      - run: bash testing/run-unit-suites.sh\n      - uses: ./local@8f4b7f8d\n",
        encoding="utf-8")
    (repo / ".github" / "required-jobs-policy.json").write_text(json.dumps({
        "required_jobs": ["unit"],
        "events": {e: {"accepted_skips": []} for e in
                   ("pull_request", "push", "schedule", "workflow_dispatch")},
    }), encoding="utf-8")
    # fixtures de la fase 4 (los cinco gates del patrón, todos en verde):
    (repo / ".github" / "needs").mkdir()
    (repo / ".github" / "needs" / "needs.json").write_text(
        json.dumps({"unit": {"result": "success"}}), encoding="utf-8")
    (repo / ".github" / "local-ci-parity-policy.json").write_text(json.dumps({
        "ci": {"workflow": ".github/workflows/ci.yml", "job": "unit"},
        "local": {"entrypoint": "testing/run-unit-suites.sh"}, "exclusions": [],
    }), encoding="utf-8")
    (repo / "testing").mkdir()
    (repo / "testing" / "run-unit-suites.sh").write_text(
        "#!/bin/sh\nbash testing/run-unit-suites.sh\n", encoding="utf-8")
    (repo / "AGENTS.md").write_text("# consumer\n", encoding="utf-8")
    (repo / ".github" / "host-labels-snapshot.json").write_text(
        json.dumps(["status:approved", "chain:partial", "status:blocked"],
                   ensure_ascii=False), encoding="utf-8")
    (repo / "ci-pattern.yaml").write_text(
        "P09_labels_type:\n"
        "  bug_report.yml: type:bug\n"
        "  feature_request.yml: type:feature\n",
        encoding="utf-8")
    (repo / ".github" / "host-contract.json").write_text(json.dumps({
        "contract_version": 1,
        "labels": [
            {"name": "status:approved", "class": "host-enforced"},
            {"name": "chain:partial", "class": "host-enforced"},
            {"name": "status:blocked", "class": "documented-only"},
        ],
        "required_checks": [
            {"name": "branch-name", "class": "host-enforced"},
            {"name": "required", "class": "host-enforced"},
        ],
        "merge_methods": {
            "allow_merge_commit": {"declared": True, "class": "host-enforced"},
            "allow_squash_merge": {"declared": False, "class": "host-enforced"},
            "allow_rebase_merge": {"declared": False, "class": "host-enforced"},
            "allow_auto_merge": {"declared": True, "class": "host-enforced"},
        },
        "protection": {
            "enforce_admins": {"declared": True, "class": "host-enforced"},
            "required_linear_history": {"declared": True, "class": "host-enforced"},
            "required_conversation_resolution": {"declared": True, "class": "documented-only"},
            "allow_force_pushes": {"declared": False, "class": "host-enforced"},
            "allow_deletions": {"declared": False, "class": "host-enforced"},
        },
        "rulesets": [{"name": "branch-protection", "enforcement": "active", "class": "host-enforced"}],
    }), encoding="utf-8")
    (repo / "api").mkdir()
    (repo / "api" / "repo.json").write_text(
        '{"allow_merge_commit": true, "allow_squash_merge": false, "allow_rebase_merge": false, "allow_auto_merge": true}\n',
        encoding="utf-8")
    (repo / "api" / "labels.json").write_text(
        '[{"name": "status:approved"}, {"name": "chain:partial"}]\n', encoding="utf-8")
    (repo / "api" / "branch-protection.json").write_text(
        '{"enforce_admins": {"enabled": true}, "required_linear_history": {"enabled": true}, '
        '"required_conversation_resolution": {"enabled": false}, "allow_force_pushes": {"enabled": false}, '
        '"allow_deletions": {"enabled": false}, "required_status_checks": {"contexts": ["branch-name", "required"], "checks": []}}\n',
        encoding="utf-8")
    (repo / "api" / "rulesets.json").write_text(
        '[{"name": "branch-protection", "enforcement": "active", "class": "host-enforced"}]\n',
        encoding="utf-8")
    # mini-skill con un gate inventariado (HR-1) bajo `assets/gate/`; sirve
    # de sujeto gobernado a los tests de cobertura de la fase 0 sin acoplarse
    # al catálogo canónico de la skill adoptada.
    (repo / "skill" / "assets" / "gate" / "tests").mkdir(parents=True)
    (repo / "skill" / "references").mkdir()
    (repo / "skill" / "SKILL.md").write_text(
        "---\nname: mini\n---\n\n- **HR-1 — regla**\n", encoding="utf-8")
    (repo / "skill" / "assets" / "gate" / "gate.py").write_text("# gate\n", encoding="utf-8")
    (repo / "skill" / "assets" / "gate" / "tests" / "test_gate.py").write_text(
        "print('ok')\n", encoding="utf-8")
    if contract is None:
        contract = {}
    if contract_raw is None:
        contract_raw = json.dumps(contract, indent=2, ensure_ascii=False)
    contract_path = repo / ".github" / "ci-pattern-adoption.json"
    contract_path.write_text(contract_raw, encoding="utf-8")
    # #312: la fase 0 exige cobertura exhaustiva — el fixture declara
    # host_labels (instantánea vacía por defecto) y cubre los ficheros
    # gobernados que él mismo crea, salvo que el test los declare ya.
    phases = contract.get("phases") if isinstance(contract, dict) else None
    p0 = phases.get("0") if isinstance(phases, dict) else None
    if isinstance(p0, dict) and "inventory" in p0:
        p0.setdefault("host_labels", ".github/host-labels-snapshot.json")
        (repo / ".github" / "host-labels-snapshot.json").write_text(
            "[]", encoding="utf-8")
        covered = {e.get(k) for e in p0["inventory"] if isinstance(e, dict)
                   for k in ("artifact", "path", "target")}
        if p0["inventory"]:
            for rel in (".github/workflows/ci.yml", ".github/host-contract.json",
                        ".github/env-isolation-allowlist.json",
                        ".github/required-jobs-policy.json",
                        ".github/local-ci-parity-policy.json", "AGENTS.md",
                        # #311: script de gobierno por REFERENCIA (el `run:`
                        # del ci.yml y el entrypoint de paridad HR-4)
                        "testing/run-unit-suites.sh"):
                if rel not in covered:
                    p0["inventory"].append(
                        {"artifact": rel, "disposition": "adopted", "target": rel})
        contract_path.write_text(
            json.dumps(contract, indent=2, ensure_ascii=False), encoding="utf-8")
    return repo, contract_path


def full_phase4_gates(extra=None):
    """Los cinco gates del patrón con entradas declaradas por el consumer."""
    gates = [
        {"gate": "host-readback",
         "args": ["--contract", "@.github/host-contract.json",
                  "--snapshot", "repo=@api/repo.json",
                  "--snapshot", "labels=@api/labels.json",
                  "--snapshot", "branch-protection=@api/branch-protection.json",
                  "--snapshot", "rulesets=@api/rulesets.json"]},
        {"gate": "required-jobs",
         "args": ["--policy", "@.github/required-jobs-policy.json",
                  "--event", "pull_request", "--needs-file", "@.github/needs/needs.json"]},
        {"gate": "local-ci-parity",
         "args": ["--repo", "@.", "--policy", "@.github/local-ci-parity-policy.json"]},
        {"gate": "workflow-policy",
         "args": ["--workflow", "@.github/workflows/ci.yml",
                  "--required-jobs", "@.github/required-jobs-policy.json"]},
    ]
    if extra:
        gates.extend(extra)
    return {"gates": gates}


def docs_phase_decl():
    return {"operating_doc": {"doc": "docs/ci-pattern-flow.md",
                              "linked_from": "AGENTS.md"}}


def seed_operating_doc(repo):
    """Genera el documento operativo en el destino y lo enlaza desde
    AGENTS.md (fase 5 en verde para los fixtures de fases posteriores)."""
    decl = docs_phase_decl()
    out, err = io.StringIO(), io.StringIO()
    old = os.environ.copy()
    os.environ["CI_PATTERN_SKILL_DIR"] = str(SKILL_DIR)
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            CLI.main(["adoption", "generate-doc", str(repo),
                      "--out", str(repo / decl["operating_doc"]["doc"])])
    finally:
        os.environ.clear()
        os.environ.update(old)
    agents = repo / "AGENTS.md"
    agents.write_text(
        agents.read_text(encoding="utf-8")
        + "Flujo operativo: docs/ci-pattern-flow.md\n", encoding="utf-8")
    # #305: el bloque canónico (índice + mandato de no invención) se copia
    # del documento generado, como haría el operador del consumer.
    doc_text = (repo / decl["operating_doc"]["doc"]).read_text(encoding="utf-8")
    begin = "<!-- ci-pattern-agents:begin -->"
    end = "<!-- ci-pattern-agents:end -->"
    agents.write_text(
        agents.read_text(encoding="utf-8")
        + doc_text[doc_text.index(begin):doc_text.index(end) + len(end)]
        + "\n", encoding="utf-8")
    # AGENTS.md es artefacto sellado (HR-46): el manifiesto se re-sella tras
    # el cambio legítimo del enlace.
    seed_manifest(repo)


def complete_contract(upto=0, **overrides):
    """Contrato válido con las fases 0..upto declaradas completas."""
    phases = {}
    if upto >= 0:
        phases["0"] = {
            "inventory": [
                {"artifact": "legacy-pipeline", "disposition": "retired"},
                {
                    "artifact": ".github/workflows/ci.yml",
                    "disposition": "adopted",
                    "target": ".github/workflows/ci.yml",
                },
                {"artifact": ".github/host-contract.json",
                 "disposition": "adopted", "target": ".github/host-contract.json"},
                {"artifact": "AGENTS.md",
                 "disposition": "adopted", "target": "AGENTS.md"},
                {"artifact": ".github/env-isolation-allowlist.json",
                 "disposition": "adopted", "target": ".github/env-isolation-allowlist.json"},
                {"artifact": ".github/required-jobs-policy.json",
                 "disposition": "adopted", "target": ".github/required-jobs-policy.json"},
                {"artifact": ".github/local-ci-parity-policy.json",
                 "disposition": "adopted", "target": ".github/local-ci-parity-policy.json"},
                {"artifact": "labels status:approved, chain:partial, status:blocked",
                 "disposition": "adopted", "target": ".github/host-contract.json"},
            ],
            "removal_confirmed": True,
            "host_labels": ".github/host-labels-snapshot.json",
        }
    if upto >= 1:
        phases["1"] = {"evidence": {"path": "docs/adoption/propagation-evidence.md"}}
    if upto >= 2:
        phases["2"] = {
            "env_isolation": {
                "root": "tests",
                "allowlist": ".github/env-isolation-allowlist.json",
            }
        }
    if upto >= 3:
        phases["3"] = {
            "evidence": {"path": "docs/adoption/governance-replacement-pr.md"},
            "pattern_paths": [".github/workflows/ci.yml"],
            "host_snapshots": {
                "branch-protection": "docs/adoption/branch-protection.json",
                "rulesets": "docs/adoption/rulesets.json",
            },
        }
    if upto >= 4:
        phases["4"] = full_phase4_gates()
    if upto >= 5:
        phases["5"] = docs_phase_decl()
    if upto >= 6:
        phases["6"] = {
            "prs": [{"number": 101, "sha": "b" * 40, "conclusion": "success"}]
        }
    phases.update(overrides.get("phases", {}))
    doc = {"schema_version": 1, "phases": phases}
    doc.update({k: v for k, v in overrides.items() if k != "phases"})
    return doc


def init_git_with_commit(repo):
    """Convierte el fixture en repo git con un commit real; devuelve su SHA."""
    subprocess.run(["git", "-C", str(repo), "init", "-q", "-b", "main"], check=True)
    subprocess.run(["git", "-C", str(repo), "config", "user.name", "fixture"], check=True)
    subprocess.run(["git", "-C", str(repo), "config", "user.email", "fixture@invalid"], check=True)
    (repo / "README.md").write_text("consumer\n", encoding="utf-8")
    subprocess.run(["git", "-C", str(repo), "add", "-A"], check=True)
    subprocess.run(["git", "-C", str(repo), "commit", "-q", "-m", "chore: base"], check=True)
    head = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"],
                          capture_output=True, text=True, check=True)
    return head.stdout.strip()


def touch(repo, rel, text="evidencia\n"):
    path = repo / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


HOST_CONTRACT = {
    "contract_version": 1,
    "labels": [
        {"name": "status:approved", "class": "host-enforced"},
        {"name": "chain:partial", "class": "host-enforced"},
        {"name": "status:blocked", "class": "documented-only"},
    ],
    "required_checks": [
        {"name": "branch-name", "class": "host-enforced"},
        {"name": "required", "class": "host-enforced"},
    ],
    "merge_methods": {
        "allow_merge_commit": {"declared": True, "class": "host-enforced"},
        "allow_squash_merge": {"declared": False, "class": "host-enforced"},
        "allow_rebase_merge": {"declared": False, "class": "host-enforced"},
        "allow_auto_merge": {"declared": True, "class": "host-enforced"},
    },
    "protection": {
        "enforce_admins": {"declared": True, "class": "host-enforced"},
        "required_linear_history": {"declared": True, "class": "host-enforced"},
        "required_conversation_resolution": {"declared": True, "class": "documented-only"},
        "allow_force_pushes": {"declared": False, "class": "host-enforced"},
        "allow_deletions": {"declared": False, "class": "host-enforced"},
    },
    "rulesets": [{"name": "branch-protection", "enforcement": "active",
                  "class": "host-enforced"}],
}

# Artefactos que el patrón instala en el consumer y por tanto fija el
# manifiesto de gobierno. make_contract_repo los crea siempre.
PATTERN_MANIFEST_FILES = (
    "AGENTS.md",
    ".github/workflows/ci.yml",
    ".github/host-contract.json",
    ".github/ci-pattern-adoption.json",
    ".github/env-isolation-allowlist.json",
    ".github/required-jobs-policy.json",
    ".github/local-ci-parity-policy.json",
)


def seed_manifest(repo, paths=PATTERN_MANIFEST_FILES):
    """Manifiesto de gobierno con el sha256 gobernado de cada artefacto
    del patrón: la identidad del patrón la fija el manifiesto verificado con
    `ci-pattern verify`, nunca el contrato de adopción (#277). Los workflows se
    sellan normalizados (#1199), así que el hash es el del contenido con los pins
    a un marcador fijo."""
    files = {}
    for rel in paths:
        files[rel] = CLI.sha256_governed(repo / rel, rel)
    touch(repo, ".governance-manifest.json",
          json.dumps({"schema_version": CLI.SCHEMA_VERSION, "files": files}, indent=2))


def seed_phase3(repo, snapshot_rulesets=None):
    """Siembra la fase 3 del fixture: instantáneas del host que casan con el
    contrato del host de make_contract_repo (cuyos rulesets son los del
    patrón, no autodeclarados) y manifiesto de gobierno."""
    touch(repo, "docs/adoption/branch-protection.json",
          json.dumps({"required_status_checks": {"enabled": True}}, indent=2))
    touch(repo, "docs/adoption/rulesets.json",
          json.dumps(snapshot_rulesets if snapshot_rulesets is not None
                     else HOST_CONTRACT["rulesets"], indent=2))
    seed_manifest(repo, list(PATTERN_MANIFEST_FILES))


def write_contract(repo, contract):
    """Reescritura del contrato tras sellar el manifiesto: re-sella, porque
    el contrato es un artefacto cubierto por el manifiesto (#277)."""
    (repo / ".github" / "ci-pattern-adoption.json").write_text(
        json.dumps(contract, indent=2, ensure_ascii=False), encoding="utf-8")
    if (repo / ".governance-manifest.json").is_file():
        seed_manifest(repo)


def seed_evidence(repo, upto=3):
    """Siembra las evidencias de las fases 1..upto y, para upto >= 3, las
    instantáneas del host de la fase 3 (#277)."""
    for rel in (
        "docs/adoption/propagation-evidence.md",
        "docs/adoption/env-isolation-green.md",
        "docs/adoption/governance-replacement-pr.md",
    )[:upto]:
        touch(repo, rel)
    if upto >= 3:
        seed_phase3(repo)


class ContractValidationTests(unittest.TestCase):
    def test_missing_contract_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "consumer"
            repo.mkdir()
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("ci-pattern-adoption.json", err)

    def test_unreadable_json_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract_raw="{ not json ]")
            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("JSON", err)

    def test_unknown_schema_version_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract={"schema_version": 99, "phases": {}})
            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("schema_version", err)

    def test_phases_not_object_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract={"schema_version": 1, "phases": [1, 2]})
            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 2)

    def test_unknown_phase_key_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract={"schema_version": 1, "phases": {"9": {}}})
            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("9", err)


class PhaseZeroTests(unittest.TestCase):
    def test_phase_zero_complete_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
            code, out, _ = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_zero_empty_inventory_fails_naming(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=0)
            contract["phases"]["0"]["inventory"] = []
            repo, _ = make_contract_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 0", err)

    def test_phase_zero_bad_disposition_fails_naming_artifact(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=0)
            contract["phases"]["0"]["inventory"][0]["disposition"] = "kept"
            repo, _ = make_contract_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy-pipeline", err)
            self.assertIn("kept", err)

    def test_phase_zero_virgin_consumer_passes(self):
        # #306: los targets 'adopted' se INSTALAN en la fase 3; la fase 0
        # solo exige que apunten a artefactos canónicos del patrón. Un
        # consumer virgen (gobierno ajeno, cero artefactos del patrón) con
        # inventario completo y removal_confirmed sale 0.
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "virgin"
            (repo / ".github" / "workflows").mkdir(parents=True)
            (repo / ".github" / "pull_request_template.md").write_text(
                "legacy\n", encoding="utf-8")
            (repo / ".github" / "workflows" / "build.yml").write_text(
                "on: push\n", encoding="utf-8")
            (repo / "ci.yaml-legacy").write_text("legacy\n", encoding="utf-8")
            contract = {
                "schema_version": 1,
                "phases": {
                    "0": {
                        "inventory": [
                            {"artifact": "legacy-pr-template",
                             "disposition": "retired",
                             "path": ".github/pull_request_template.md"},
                            {"artifact": "legacy-ci-yaml",
                             "disposition": "retired", "path": "ci.yaml-legacy"},
                            {"artifact": "legacy-build-workflow",
                             "disposition": "retired",
                             "path": ".github/workflows/build.yml"},
                            {"artifact": "legacy-build-workflow-2",
                             "disposition": "adopted",
                             "target": ".github/workflows/ci.yml"},
                            {"artifact": "docs/branch-protection.md",
                             "disposition": "adopted",
                             "target": ".github/host-contract.json"},
                        ],
                        "removal_confirmed": True,
                        "host_labels": ".github/host-labels-snapshot.json",
                    }
                },
            }
            (repo / ".github" / "host-labels-snapshot.json").write_text(
                "[]", encoding="utf-8")
            (repo / ".github" / "ci-pattern-adoption.json").write_text(
                json.dumps(contract), encoding="utf-8")
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_zero_non_canonical_target_is_finding(self):
        # #306: un target 'adopted' fuera del catálogo canónico del patrón
        # es un hallazgo que lo nombra (la fase 3 instala los canónicos).
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=0)
            contract["phases"]["0"]["inventory"][1]["target"] = "Jenkinsfile"
            repo, _ = make_contract_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("canónico", err)
            self.assertIn("Jenkinsfile", err)

    def test_phase_zero_unconfirmed_removal_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=0)
            contract["phases"]["0"]["removal_confirmed"] = False
            repo, _ = make_contract_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("removal_confirmed", err)





CADETE_SKILLS = (
    "cadete-workflow", "cadete-devops", "documentation-alan-style",
    "feature-acceptance-uat", "gentle-ai-ai-slop-discipline",
    "intake-roadmap-loop", "repository-delivery-governance",
)


def make_cadete_like_repo(tmp, contract=None, with_agents=True):
    """Fixture tipo Cadete (#311): los artefactos de gobierno reales de un
    consumer con scripts de gate consumer-side, manifiesto y destino de
    propagación de skills (flota) y gobierno ajeno."""
    repo = Path(tmp) / "cadete-like"
    (repo / ".github" / "workflows").mkdir(parents=True)
    (repo / ".github" / "ISSUE_TEMPLATE").mkdir(parents=True)
    (repo / "scripts").mkdir()
    (repo / "AGENTS.md").write_text("# consumer\n", encoding="utf-8")
    (repo / ".github" / "workflows" / "ci.yml").write_text(
        "on: push\njobs:\n  ratchet:\n    steps:\n"
        "      - run: bash scripts/coverage-ratchet.sh\n", encoding="utf-8")
    (repo / ".github" / "workflows" / "ci_sdd.yml").write_text("on: push\n", encoding="utf-8")
    (repo / ".github" / "ISSUE_TEMPLATE" / "issue-canonical.yml").write_text("name: x\n", encoding="utf-8")
    (repo / ".github" / "ISSUE_TEMPLATE" / "config.yml").write_text("blank_issues_enabled: false\n", encoding="utf-8")
    (repo / ".github" / "PULL_REQUEST_TEMPLATE.md").write_text("tpl\n", encoding="utf-8")
    (repo / ".github" / "coverage-baseline.json").write_text("{}\n", encoding="utf-8")
    (repo / "scripts" / "coverage-ratchet.sh").write_text("#!/bin/sh\n", encoding="utf-8")
    (repo / "scripts" / "local-entrypoint.sh").write_text("#!/bin/sh\n", encoding="utf-8")
    (repo / "scripts" / "refresh-team-skills.ps1").write_text("# sync\n", encoding="utf-8")
    # script OPERATIVO de la APP: nadie del gobierno lo invoca (#311)
    (repo / "scripts" / "backup.ps1").write_text("# dump diario\n", encoding="utf-8")
    if with_agents:
        for name in CADETE_SKILLS:
            d = repo / ".agents" / "skills" / name
            d.mkdir(parents=True, exist_ok=True)
            (d / "SKILL.md").write_text("---\nname: x\n---\n", encoding="utf-8")
        (repo / ".team-skills.yaml").write_text(
            "skills:\n" + "".join(f"  - {n}\n" for n in CADETE_SKILLS), encoding="utf-8")
    contract.setdefault("phases", {}).setdefault("0", {}).setdefault(
        "host_labels", ".github/host-labels-snapshot.json")
    (repo / ".github" / "host-labels-snapshot.json").write_text(
        json.dumps(CADETE_HOST_LABELS, ensure_ascii=False), encoding="utf-8")
    (repo / ".github" / "ci-pattern-adoption.json").write_text(
        json.dumps(contract, indent=2, ensure_ascii=False), encoding="utf-8")
    return repo


def cadete_phase0_contract():
    """Las entradas reales de la fase 0 de Cadete (35 -> resumidas a las que
    el catálogo debe aceptar: adopted consumer-side y retired)."""
    return {
        "schema_version": 1,
        "phases": {
            "0": {
                "inventory": [
                    {"artifact": ".github/workflows/ci.yml", "disposition": "adopted",
                     "target": ".github/workflows/ci.yml"},
                    {"artifact": ".github/ISSUE_TEMPLATE/issue-canonical.yml",
                     "disposition": "adopted", "target": ".github/ISSUE_TEMPLATE/issue-canonical.yml"},
                    {"artifact": ".github/ISSUE_TEMPLATE/config.yml",
                     "disposition": "adopted", "target": ".github/ISSUE_TEMPLATE/config.yml"},
                    {"artifact": ".github/PULL_REQUEST_TEMPLATE.md",
                     "disposition": "adopted", "target": ".github/PULL_REQUEST_TEMPLATE.md"},
                    {"artifact": ".github/coverage-baseline.json",
                     "disposition": "adopted", "target": ".github/coverage-baseline.json"},
                    {"artifact": "AGENTS.md", "disposition": "adopted", "target": "AGENTS.md"},
                    {"artifact": "scripts/coverage-ratchet.sh (ratchet HR-15/16)",
                     "disposition": "adopted", "target": "scripts/coverage-ratchet.sh"},
                    {"artifact": ".team-skills.yaml (manifiesto de la flota)",
                     "disposition": "adopted", "target": ".team-skills.yaml"},
                    {"artifact": ".agents/skills/ (destino de la flota)",
                     "disposition": "adopted", "target": ".agents/skills/"},
                    {"artifact": ".github/workflows/ci_sdd.yml (staging)",
                     "disposition": "retired", "path": ".github/workflows/ci_sdd.yml"},
                    {"artifact": "scripts/refresh-team-skills.ps1 (sync manual)",
                     "disposition": "retired", "path": "scripts/refresh-team-skills.ps1"},
                    {"artifact": ".agents/skills/cadete-workflow/ (flujo de revisor único)",
                     "disposition": "retired", "path": ".agents/skills/cadete-workflow"},
                    {"artifact": "label:borrador", "disposition": "retired"},
                    {"artifact": "label:en-progreso", "disposition": "retired"},
                    {"artifact": "label:pendiente-revisión", "disposition": "retired"},
                    {"artifact": "label:aprobado", "disposition": "retired"},
                    {"artifact": "label:rechazado", "disposition": "retired"},
                    {"artifact": "label:mergeado", "disposition": "retired"},
                    {"artifact": "label:abierto", "disposition": "retired"},
                    {"artifact": "label:auto-aprobado (flujo de revisor único)",
                     "disposition": "retired"},
                    {"artifact": "label:needs-review (flujo de revisor único)",
                     "disposition": "retired"},
                    {"artifact": "label:security", "disposition": "retired"},
                    {"artifact": "label:prioridad:alta", "disposition": "retired"},
                    {"artifact": "label:prioridad:media", "disposition": "retired"},
                    {"artifact": "label:prioridad:baja", "disposition": "retired"},
                    {"artifact": "label:priority:high", "disposition": "retired"},
                    {"artifact": "label:priority:critical", "disposition": "retired"},
                    {"artifact": "label:salud-código", "disposition": "retired"},
                    {"artifact": "label:size:exception", "disposition": "retired"},
                    {"artifact": "label:type:mantenimiento", "disposition": "retired"},
                    {"artifact": "labels type:bug, type:docs, type:feature, "
                                 "type:chore, status:approved (del patrón)",
                     "disposition": "adopted", "target": ".github/host-contract.json"},
                    {"artifact": "host: métodos de merge y protección (GET solo lectura)",
                     "disposition": "adopted", "target": ".github/host-contract.json"},
                ],
                "removal_confirmed": True,
            }
        },
    }


class Phase0CatalogTests(unittest.TestCase):
    """#311: el catálogo canónico de la fase 0 es DATO declarado por el
    patrón y representa los artefactos consumer-side (scripts de gate,
    manifiesto y destino de la propagación de skills)."""

    def test_phase_zero_gate_scripts_are_canonical(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_cadete_like_repo(tmp, contract=cadete_phase0_contract())
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_zero_fleet_channel_is_canonical(self):
        # mismo fixture, sin .agents/.team-skills: solo el canal de flota
        # declarado adopted no puede ser hallazgo por ubicación
        with tempfile.TemporaryDirectory() as tmp:
            contract = cadete_phase0_contract()
            repo = make_cadete_like_repo(tmp, contract=contract)
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_three_vendored_skills_covered_by_directory_entry(self):
        # la fase 3 escanea las rutas del catálogo (ahora también
        # .agents/skills/*): un entry adopted de directorio las cubre
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            contract["phases"]["0"]["inventory"].append(
                {"artifact": ".agents/skills/ (destino de la flota)",
                 "disposition": "adopted", "target": ".agents/skills/"})
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_phase3(repo)
            for name in CADETE_SKILLS:
                d = repo / ".agents" / "skills" / name
                d.mkdir(parents=True, exist_ok=True)
                (d / "SKILL.md").write_text("---\nname: x\n---\n", encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, err)
            self.assertNotIn("no inventariado", err)



CADETE_HOST_LABELS = [
    "type:feature", "type:docs", "type:bug", "aprobado", "prioridad:alta",
    "prioridad:baja", "prioridad:media", "rechazado", "salud-código",
    "auto-aprobado", "type:mantenimiento", "borrador", "en-progreso",
    "pendiente-revisión", "mergeado", "abierto", "needs-review", "security",
    "priority:high", "size:exception", "type:chore", "priority:critical",
    "status:approved",
]


class Phase0ExhaustivenessTests(unittest.TestCase):
    """#312: la fase 0 EJECUTA la cobertura del inventario — cada artefacto
    de gobierno escaneado (rutas del catálogo, etiquetas del host vía
    instantánea declarada, sin red) exige su entrada; una omisión es
    hallazgo nombrándolo."""

    def cadete_contract(self, omit=None):
        contract = cadete_phase0_contract()
        contract["phases"]["0"]["host_labels"] = ".github/host-labels-snapshot.json"
        inv = contract["phases"]["0"]["inventory"]
        if omit:
            inv[:] = [e for e in inv if not any(o in e.get("artifact", "") for o in omit)]
        return contract

    def write_labels_snapshot(self, repo, labels=CADETE_HOST_LABELS):
        (repo / ".github" / "host-labels-snapshot.json").write_text(
            json.dumps(labels, ensure_ascii=False), encoding="utf-8")

    def test_phase_zero_requires_host_labels_declaration(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = cadete_phase0_contract()  # sin host_labels
            repo = make_cadete_like_repo(tmp, contract=contract)
            contract["phases"]["0"].pop("host_labels", None)
            (repo / ".github" / "ci-pattern-adoption.json").write_text(
                json.dumps(contract, indent=2, ensure_ascii=False), encoding="utf-8")
            self.write_labels_snapshot(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("host_labels", err)

    def test_phase_zero_omitted_label_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.cadete_contract(omit=["label:security"])
            repo = make_cadete_like_repo(tmp, contract=contract)
            self.write_labels_snapshot(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("security", err)

    def test_phase_zero_all_labels_covered_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.cadete_contract()
            repo = make_cadete_like_repo(tmp, contract=contract)
            self.write_labels_snapshot(repo)
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_zero_omitted_tree_artifact_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.cadete_contract(omit=[".github/workflows/ci.yml"])
            repo = make_cadete_like_repo(tmp, contract=contract)
            self.write_labels_snapshot(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn(".github/workflows/ci.yml", err)

    def test_phase_zero_gate_inputs_are_self_excluded(self):
        # el propio contrato y la instantánea de etiquetas no exigen entrada
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.cadete_contract()
            repo = make_cadete_like_repo(tmp, contract=contract)
            self.write_labels_snapshot(repo)
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)
            self.assertNotIn("ci-pattern-adoption.json", err)
            self.assertNotIn("host-labels-snapshot", err)

    def test_phase_zero_unreadable_labels_snapshot_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.cadete_contract()
            repo = make_cadete_like_repo(tmp, contract=contract)
            (repo / ".github" / "host-labels-snapshot.json").write_text("{no", encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 2, err)

class ScriptGovernanceByReferenceTests(unittest.TestCase):
    """#311: un script es GOBIERNO si el gobierno lo invoca, no por su
    ubicación. El glob `scripts/*` metía en el inventario (y por tanto en la
    lista de retirada) scripts operativos de la APP: en Cadete,
    `scripts/importar-backup-mysql-prod.ps1` y `scripts/local-entrypoint.sh`
    (el entrypoint de Docker, invocado por Dockerfile.local, no por el CI)."""

    def test_unreferenced_app_scripts_are_not_governed(self):
        # backup.ps1 y local-entrypoint.sh están en el árbol y NO tienen
        # entrada: no son gobierno y la fase 0 no los nombra.
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_cadete_like_repo(tmp, contract=cadete_phase0_contract())
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)
            self.assertNotIn("backup.ps1", out + err)
            self.assertNotIn("local-entrypoint.sh", out + err)

    def test_script_referenced_by_workflow_must_be_inventoried(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_cadete_like_repo(tmp, contract=cadete_phase0_contract())
            (repo / "scripts" / "ci-check.sh").write_text("#!/bin/sh\n", encoding="utf-8")
            wf = repo / ".github" / "workflows" / "ci.yml"
            wf.write_text(wf.read_text(encoding="utf-8")
                          + "      - run: bash scripts/ci-check.sh\n", encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("scripts/ci-check.sh", err)

    def test_referenced_script_is_canonical_once_inventoried(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = cadete_phase0_contract()
            contract["phases"]["0"]["inventory"].append({
                "artifact": "scripts/ci-check.sh (gate consumer-side)",
                "disposition": "adopted", "target": "scripts/ci-check.sh"})
            repo = make_cadete_like_repo(tmp, contract=contract)
            (repo / "scripts" / "ci-check.sh").write_text("#!/bin/sh\n", encoding="utf-8")
            wf = repo / ".github" / "workflows" / "ci.yml"
            wf.write_text(wf.read_text(encoding="utf-8")
                          + "      - run: sh scripts/ci-check.sh\n", encoding="utf-8")
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_contract_declared_governed_path_is_canonical(self):
        # la vía explícita: el consumer declara en su contrato el script de
        # gobierno que ninguna fuente del patrón referencia.
        with tempfile.TemporaryDirectory() as tmp:
            contract = cadete_phase0_contract()
            contract["phases"]["0"]["declared_governed_paths"] = ["scripts/backup.ps1"]
            contract["phases"]["0"]["inventory"].append({
                "artifact": "scripts/backup.ps1 (declarado por el consumer)",
                "disposition": "adopted", "target": "scripts/backup.ps1"})
            repo = make_cadete_like_repo(tmp, contract=contract)
            code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_undeclared_app_script_target_is_not_canonical(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = cadete_phase0_contract()
            contract["phases"]["0"]["inventory"].append({
                "artifact": "scripts/backup.ps1 (no es gobierno)",
                "disposition": "adopted", "target": "scripts/backup.ps1"})
            repo = make_cadete_like_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("scripts/backup.ps1", err)
            self.assertIn("canónico", err)


class SequencingTests(unittest.TestCase):
    def test_earlier_phase_failure_blocks_later_phase(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=1)
            touch(Path(tmp), "docs/adoption/propagation-evidence.md")
            repo, _ = make_contract_repo(tmp, contract=contract)
            # la fase 0 está rota: inventory vacío
            contract["phases"]["0"]["inventory"] = []
            (repo / ".github" / "ci-pattern-adoption.json").write_text(
                json.dumps(contract), encoding="utf-8"
            )
            code, _, err = run_cli(["adoption", "check", "--phase", "1", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 0", err)

    def test_phase_1_with_evidence_missing_file_fails_naming(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=1)
            repo, _ = make_contract_repo(tmp, contract=contract)
            code, _, err = run_cli(["adoption", "check", "--phase", "1", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("docs/adoption/propagation-evidence.md", err)

    def test_phase_1_with_evidence_present_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=1)
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            code, out, _ = run_cli(["adoption", "check", "--phase", "1", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_2_with_evidence_present_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            touch(repo, "docs/adoption/env-isolation-green.md")
            code, out, _ = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 0, out)


class Phase3SubstitutionTests(unittest.TestCase):
    """#277: la fase 3 ejecuta la comprobación sobre el árbol y las
    instantáneas del host; no acepta un fichero de evidencia."""

    def test_residual_foreign_governance_fails_naming_each(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/workflows/legacy-deploy.yml", text="on: push\n")
            touch(repo, ".github/pull_request_template.md", text="legacy\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy-deploy.yml", err)
            self.assertIn("pull_request_template.md", err)

    def test_phase_three_retired_path_still_present_fails(self):
        # #306: la eliminación del gobierno ajeno ocurre en la fase 3; un
        # retired declarado con 'path' que SIGUE en el árbol es residual.
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            contract["phases"]["0"]["inventory"].append(
                {"artifact": "legacy-pr-template", "disposition": "retired",
                 "path": ".github/pull_request_template.md"}
            )
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/pull_request_template.md", text="legacy\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy-pr-template", err)
            self.assertIn(".github/pull_request_template.md", err)

    def test_full_substitution_is_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, out, _ = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, out + _)

    def test_retired_artifact_still_in_tree_is_residual(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3, phases={
                "0": {"inventory": [
                    {"artifact": ".github/workflows/legacy.yml",
                     "disposition": "retired"},
                ], "removal_confirmed": True},
            })
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/workflows/legacy.yml", text="on: push\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy.yml", err)
            self.assertIn("residual", err)

    def test_host_ruleset_residual_fails_naming_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo, snapshot_rulesets=[
                {"name": "branch-protection", "enforcement": "active", "class": "branch"},
                {"name": "legacy-apap", "enforcement": "active", "class": "branch"},
            ])
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy-apap", err)

    def test_retired_host_ruleset_still_active_is_residual(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3, phases={
                "0": {"inventory": [
                    {"artifact": "ruleset:legacy-apap", "disposition": "retired"},
                ], "removal_confirmed": True},
            })
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo, snapshot_rulesets=[
                {"name": "legacy-apap", "enforcement": "active", "class": "branch"}])
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy-apap", err)

    def test_legacy_workflow_claimed_as_pattern_path_fails(self):
        """La identidad del patrón no es autodeclarada: un workflow legado
        listado en pattern_paths pero ausente del manifiesto es residual."""
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            contract["phases"]["3"]["pattern_paths"].append(
                ".github/workflows/legacy.yml")
            contract["phases"]["0"]["inventory"].append(
                {"artifact": ".github/workflows/legacy.yml",
                 "disposition": "retired", "path": ".github/workflows/legacy.yml"})
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/workflows/legacy.yml", text="on: push\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("legacy.yml", err)
            self.assertIn("manifiesto", err)

    def test_modified_pattern_file_fails_via_verify(self):
        """Un artefacto del patrón modificado respecto del sha256 del
        manifiesto falla aunque pattern_paths lo declare."""
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/workflows/ci.yml", text="on: push\npirata: true\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("ci.yml", err)

    def test_untracked_governed_paths_are_residual(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, ".github/CODEOWNERS", text="* @legacy\n")
            touch(repo, ".github/dependabot.yml", text="version: 2\n")
            touch(repo, ".pre-commit-config.yaml", text="repos: []\n")
            touch(repo, ".husky/pre-commit", text="echo legacy\n")
            touch(repo, "lefthook.yml", text="pre-commit:\n")
            touch(repo, ".github/legacy-policy.json", text="{}\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            for named in ("CODEOWNERS", "dependabot.yml", "pre-commit-config.yaml",
                          "husky", "lefthook.yml", "legacy-policy.json"):
                self.assertIn(named, err)

    def test_missing_manifest_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            (repo / ".governance-manifest.json").unlink()
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("governance-manifest", err)

    def test_missing_host_contract_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            (repo / ".github/host-contract.json").unlink()
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("host-contract", err)

    def test_missing_snapshot_file_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            # La fase 3 exige manifiesto y contrato del host válidos antes de
            # leer instantáneas; el objetivo del test es la snapshot ausente.
            seed_phase3(repo)
            (repo / "docs/adoption/branch-protection.json").unlink()
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("branch-protection.json", err)

    def test_malformed_ruleset_snapshot_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 2)
            seed_phase3(repo)
            touch(repo, "docs/adoption/rulesets.json",
                  text="[{\"enforcement\": \"active\"}]\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("name", err)


class Phase3WiringTests(unittest.TestCase):
    """#1027 (HR-72): la fase 3 se mide en un pull request.

    El control que vigila el cableado de los gates era el unico sin cablear:
    la fase 3 comprueba que cada job de `P23_required_jobs` este en los
    workflows, y ningun PR la ejecutaba. Con `governance_seals` declarado, el
    propio `adoption check --phase 3` —invocado como comando de primer nivel
    desde el job de PR— exige que ese job exista y figure en `required_jobs`.
    """

    def _repo(self, tmp, wired):
        contract = complete_contract(upto=3)
        contract["governance_seals"] = {"job": "governance-seals"}
        repo, _ = make_contract_repo(tmp, contract=contract)
        seed_evidence(repo, 3)
        if wired:
            # Antes de sellar: ci.yml y la politica son artefactos del manifiesto
            # (#277), asi que escribirlos despues los dejaria fuera de su sello.
            ci = repo / ".github/workflows/ci.yml"
            ci.write_text(ci.read_text(encoding="utf-8")
                          + "  governance-seals:\n    runs-on: ubuntu-latest\n",
                          encoding="utf-8")
            policy = repo / ".github/required-jobs-policy.json"
            data = json.loads(policy.read_text(encoding="utf-8"))
            data["required_jobs"].append("governance-seals")
            policy.write_text(json.dumps(data), encoding="utf-8")
        seed_phase3(repo)
        return repo

    def test_job_declarado_y_sin_cablear_rojo_en_el_pr_que_lo_declara(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, wired=False)
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("HR-72", err)
            self.assertIn("governance-seals", err)

    def test_job_cableado_y_exigido_da_verde(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, wired=True)
            code, out, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_sin_declarar_el_job_la_fase_3_sigue_en_verde(self):
        # Fail-open: la adopcion previa de un consumer que no declara no se
        # rompe por la aparicion de esta HR.
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=3))
            seed_evidence(repo, 3)
            seed_phase3(repo)
            code, out, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, out + err)


class LaterPhasesTests(unittest.TestCase):
    def test_phase_6_sha_must_exist_in_history(self):
        # auditoría de contenido: un sha declarado que no existe en el
        # historial del repo es una aceptación fabricada — hallazgo.
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            repo, _ = make_contract_repo(tmp, contract=contract)
            init_git_with_commit(repo)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)

            code, _, err = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("historial", err)

    def test_phase_6_sha_from_real_history_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            repo, _ = make_contract_repo(tmp, contract=contract)
            sha = init_git_with_commit(repo)
            contract["phases"]["6"] = {
                "prs": [{"number": 101, "sha": sha, "conclusion": "success"}]
            }
            write_contract(repo, contract)
            seed_operating_doc(repo)
            seed_evidence(repo, 3)
            code, out, _ = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_6_non_git_repo_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_operating_doc(repo)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("git", err)

    def test_phase_3_missing_declaration_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            touch(repo, "docs/adoption/env-isolation-green.md")
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("sin declaración", err)

    def test_phase_6_acceptance_prs_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            repo, _ = make_contract_repo(tmp, contract=contract)
            sha = init_git_with_commit(repo)
            contract["phases"]["6"] = {
                "prs": [{"number": 101, "sha": sha, "conclusion": "success"}]
            }
            write_contract(repo, contract)
            seed_operating_doc(repo)
            seed_evidence(repo, 3)
            code, out, _ = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_6_bad_sha_fails_naming(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            contract["phases"]["6"] = {
                "prs": [{"number": 12, "sha": "abc123", "conclusion": "success"}]
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)

            code, _, err = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("12", err)

    def test_phase_6_not_green_conclusion_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            contract["phases"]["6"] = {
                "prs": [{"number": 12, "sha": "a" * 40, "conclusion": "failure"}]
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)

            code, _, err = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("failure", err)

    def test_phase_6_empty_prs_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            contract["phases"]["6"] = {"prs": []}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)

            code, _, err = run_cli(["adoption", "check", "--phase", "6", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 6", err)

    def test_all_green_contract_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_operating_doc(repo)
            seed_evidence(repo, 3)
            sha = init_git_with_commit(repo)
            contract["phases"]["6"] = {
                "prs": [{"number": 101, "sha": sha, "conclusion": "success"}]
            }
            write_contract(repo, contract)
            # El contrato reescrito es un artefacto cubierto por el
            # manifiesto: se re-sella tras el cambio legítimo.
            seed_manifest(repo)
            code, out, _ = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 0, out)

    def test_all_with_failing_late_phase_names_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=6)
            contract["phases"]["6"] = {"prs": []}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)

            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 6", err)


class ExecutionGateTests(unittest.TestCase):
    """Auditoría de contenido: las fases 2 y 4 EJECUTAN los gates del patrón
    sobre el consumer (subprocess, stdlib, solo lectura); una declaración
    con un 'ok' en un .md ya no pasa."""

    @staticmethod
    def write_allowlist(repo, env_vars=None):
        path = repo / ".github" / "env-isolation-allowlist.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"env_vars": env_vars or []}), encoding="utf-8")
        return path

    def test_phase_2_env_isolation_green_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            contract["phases"]["2"] = {
                "env_isolation": {
                    "root": "tests",
                    "allowlist": ".github/env-isolation-allowlist.json",
                }
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            code, out, _ = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_2_env_leak_is_finding_naming_var(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            contract["phases"]["2"] = {
                "env_isolation": {
                    "root": "tests",
                    "allowlist": ".github/env-isolation-allowlist.json",
                }
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            touch(repo, "tests/test_smoke.sh", text='#!/bin/sh\necho "$DEPLOY_TOKEN"\n')
            code, _, err = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("DEPLOY_TOKEN", err)

    def test_phase_2_missing_allowlist_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            contract["phases"]["2"] = {
                "env_isolation": {"root": "tests", "allowlist": ".github/no.json"}
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            code, _, _ = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 2)

    def test_phase_2_missing_root_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            contract["phases"]["2"] = {
                "env_isolation": {
                    "root": "no-such-tests",
                    "allowlist": ".github/env-isolation-allowlist.json",
                }
            }
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/propagation-evidence.md")
            code, _, _ = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 2)

    def test_phase_2_without_env_isolation_key_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=2)
            contract["phases"]["2"] = {"evidence": {"path": "docs/adoption/env-isolation-green.md"}}
            repo, _ = make_contract_repo(tmp, contract=contract)
            touch(repo, "docs/adoption/env-isolation-green.md")
            touch(repo, "docs/adoption/propagation-evidence.md")
            code, _, err = run_cli(["adoption", "check", "--phase", "2", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("env_isolation", err)

    @staticmethod
    def phase4_gates(workflow_text=None, extra=None):
        # auditoría: la fase 4 exige EXACTAMENTE el conjunto de gates del
        # patrón — el consumer elige las ENTRADAS, no el subconjunto.
        return full_phase4_gates(extra=extra)

    def test_phase_4_missing_gate_is_finding_naming_it(self):
        # el consumer declara UNO en verde: faltan los otros cuatro
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = {"gates": [
                {"gate": "workflow-policy",
                 "args": ["--workflow", "@.github/workflows/ci.yml",
                          "--required-jobs", "@.github/required-jobs-policy.json"]},
            ]}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1)
            for expected in ("host-readback", "required-jobs", "local-ci-parity"):
                self.assertIn(expected, err)

    def test_phase_4_duplicate_gate_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = self.phase4_gates(
                extra=[{"gate": "workflow-policy",
                        "args": ["--workflow", "@.github/workflows/ci.yml",
                                 "--required-jobs", "@.github/required-jobs-policy.json"]}])
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("duplicado", err)

    def test_phase_4_executed_gates_green_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = self.phase4_gates(None)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, out, _ = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 0, out)

    def test_phase_4_failing_gate_is_finding_naming_gate(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = self.phase4_gates(None)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            touch(repo, ".github/workflows/ci.yml",
                  text="on:\n  pull_request:\njobs:\n  build:\n    steps:\n      - uses: actions/checkout@v4\n")
            # El workflow reescrito es un artefacto cubierto por el
            # manifiesto: se re-sella tras el cambio legítimo del escenario.
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("workflow-policy", err)

    def test_phase_4_unknown_gate_name_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = {"gates": [{"gate": "ghost-gate", "args": []}]}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, _ = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 2)

    def test_phase_4_missing_input_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = full_phase4_gates()
            # una entrada del workflow-policy apunta a un fichero ausente
            gates = contract["phases"]["4"]["gates"]
            [g for g in gates if g["gate"] == "workflow-policy"][0]["args"][1] = \
                "@.github/workflows/missing.yml"
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("entrada ausente", err)

    def test_phase_4_empty_gates_list_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = {"gates": []}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 4", err)

    def test_phase_4_contracts_key_is_no_longer_accepted(self):
        # la forma anterior (existencia + parseo) es el defecto corregido
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["4"] = {"contracts": [".github/host-contract.json"]}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("gates", err)




def drift_operating_doc(repo, old, new):
    """Introduce una deriva puntual en el documento operativo del destino."""
    doc = repo / "docs" / "ci-pattern-flow.md"
    text = doc.read_text(encoding="utf-8")
    assert old in text, f"fixture de deriva inválido: {old!r} no está en el doc"
    doc.write_text(text.replace(old, new, 1), encoding="utf-8")


class AgentsStructureTests(unittest.TestCase):
    """HR-55 (#341): la activacion depende de la DECLARACION agents_md en el
    contrato de adopción, no de la presencia del bloque (fail-closed)."""

    def _fixture(self, root, agents, manifest_sha=None, declared=True):
        repo = Path(root)
        (repo / "AGENTS.md").write_text(agents, encoding="utf-8")
        contract = {}
        if manifest_sha is not None:
            (repo / ".team-skills.yaml").write_text(
                f"slice_block_sha256: {manifest_sha}\n", encoding="utf-8")
        if declared:
            contract["agents_md"] = "propagated"
        return repo, contract

    def _run(self, repo, contract, expect_findings):
        findings = CLI._check_agents_structure(repo, contract)
        self.assertEqual(bool(findings), expect_findings,
                         f"findings={findings}")
        return findings

    def test_conforme_da_verde(self):
        body = "cuerpo del bloque\nlinea 2"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Titulo del repo\n\n"
            "<!-- personal-skills:slice:web @ v0000000 -->\n" + body +
            "\n<!-- /personal-skills:slice:web -->\n\n"
            "## Local\n\n- Leer `docs/reglas.md`\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            (repo / "docs").mkdir()
            (repo / "docs" / "reglas.md").write_text("x\n", encoding="utf-8")
            self._run(repo, contract, False)

    def test_sin_declaracion_pasa_e_informa(self):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, "# Titulo\n", declared=False)
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self._run(repo, contract, False)
            self.assertIn("no declarado", buf.getvalue())

    def test_declarado_sin_bloque_falla(self):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, "# Titulo\n")
            self._run(repo, contract, True)

    def test_texto_antes_del_bloque(self):
        body = "b"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Titulo\n\nIntroduccion manuscrita.\n\n"
            "<!-- personal-skills:slice:web @ v0000000 -->\n" + body +
            "\n<!-- /personal-skills:slice:web -->\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            findings = self._run(repo, contract, True)
            self.assertTrue(any("antes del bloque" in f for f in findings),
                            findings)

    def test_texto_entre_cierre_y_seccion(self):
        body = "b"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Titulo\n\n"
            "<!-- personal-skills:slice:web @ v0000000 -->\n" + body +
            "\n<!-- /personal-skills:slice:web -->\n\n"
            "comentario suelto\n\n## Local\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            findings = self._run(repo, contract, True)
            self.assertTrue(any("seccion local" in f for f in findings),
                            findings)

    def test_sha_alterado(self):
        body = "b"
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d,
                "# Titulo\n\n<!-- personal-skills:slice:web @ v0000000 -->\n"
                + body + "\n<!-- /personal-skills:slice:web -->\n",
                "a" * 64)
            findings = self._run(repo, contract, True)
            self.assertTrue(any("sha" in f for f in findings), findings)

    def test_seccion_local_duplicada(self):
        body = "b"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Titulo\n\n<!-- personal-skills:slice:web @ v0000000 -->\n"
            + body + "\n<!-- /personal-skills:slice:web -->\n\n"
            "## Local uno\n\nx\n\n## Local dos\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            findings = self._run(repo, contract, True)
            self.assertTrue(any("seccion local" in f for f in findings),
                            findings)

    def test_enlace_roto_en_seccion_local(self):
        body = "b"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Titulo\n\n<!-- personal-skills:slice:web @ v0000000 -->\n"
            + body + "\n<!-- /personal-skills:slice:web -->\n\n"
            "## Local\n\n- Leer `docs/no-existe.md`\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            findings = self._run(repo, contract, True)
            self.assertTrue(any("docs/no-existe.md" in f for f in findings),
                            findings)

    def test_english_agents_doc_is_language_neutral(self):
        # El gate juzga estructura (marcadores, #, ##, rutas), no idioma:
        # un AGENTS.md en inglés conforme pasa igual.
        body = "block body"
        sha = hashlib.sha256(body.encode("utf-8")).hexdigest()
        agents = (
            "# Repo title\n\n"
            "<!-- personal-skills:slice:web @ v0000000 -->\n" + body +
            "\n<!-- /personal-skills:slice:web -->\n\n"
            "## Local rules\n\n- Read `docs/reglas.md`\n")
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d, agents, sha)
            (repo / "docs").mkdir()
            (repo / "docs" / "reglas.md").write_text("x\n", encoding="utf-8")
            self._run(repo, contract, False)

    def test_declarado_bloque_sin_sha_de_manifiesto(self):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = self._fixture(d,
                "# Titulo\n<!-- personal-skills:slice:web @ v0000000 -->\nb\n"
                "<!-- /personal-skills:slice:web -->\n")
            findings = self._run(repo, contract, True)
            self.assertTrue(any("slice_block_sha256" in f for f in findings),
                            findings)


class SliceContentTests(unittest.TestCase):
    """HR-56 (#358): lo que compone la propagación no lleva nombres
    propios, rutas personales ni rutas personal/<usuario>/."""

    def _repo(self, root, maintainers=("mantener-one",)):
        repo = Path(root)
        reg = {"schema_version": 1,
               "repositories": [], "local_layout": {"local_root": str(repo)}}
        if maintainers is not None:
            reg["maintainers"] = list(maintainers)
        (repo / "fleet").mkdir(parents=True, exist_ok=True)
        (repo / "fleet" / "registry.json").write_text(
            json.dumps(reg), encoding="utf-8")
        return repo

    def _run(self, repo, relpath, text):
        target = repo / relpath
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
        return CLI._check_slice_content(repo, [target])

    def test_partial_con_login_de_mantenedor(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            findings = self._run(repo, "slices/partials/gov.md",
                                 "Creado por mantener-one para la flota.\n")
            self.assertTrue(any("mantener-one" in f for f in findings),
                            findings)

    def test_partial_con_ruta_personal(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            findings = self._run(
                repo, "slices/partials/gov.md",
                "Ver `personal/ardelperal/ci-pattern/SKILL.md`.\n")
            self.assertTrue(any("personal/" in f for f in findings), findings)

    def test_fragmento_con_ruta_de_usuario(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            findings = self._run(repo, "slices/dysflow/AGENTS.fragment.md",
                                 "Catalogo: C:\\Users\\adm1\\personal-skills.\n")
            self.assertTrue(any("Users" in f for f in findings), findings)

    def test_login_de_mantenedor_sgarcia(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d, maintainers=("SGarciaMontalvo",))
            findings = self._run(repo, "slices/partials/gov.md",
                                 "Revisado por SGarciaMontalvo.\n")
            self.assertTrue(any("SGarciaMontalvo" in f for f in findings),
                            findings)

    def test_marcador_mantenimiento_huerfano(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            text = ("<!-- maintenance:begin -->\n"
                    "texto de catalogo\n"
                    "Contenido propagado.\n")
            findings = self._run(repo, "slices/partials/gov.md", text)
            self.assertTrue(any("desbalanceado" in f for f in findings),
                            findings)

    def test_conforme_da_cero(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            findings = self._run(repo, "slices/partials/web.md",
                                 "Rutas instaladas en `.agents/skills/<n>/SKILL.md`.\n")
            self.assertEqual(findings, [])

    def test_region_de_mantenimiento_excluida(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d)
            text = ("<!-- maintenance:begin -->\n"
                    "Ver personal/ardelperal/ci-pattern/SKILL.md\n"
                    "<!-- maintenance:end -->\n"
                    "Contenido propagado.\n")
            findings = self._run(repo, "slices/partials/gov.md", text)
            self.assertEqual(findings, [])

    def test_sin_lista_de_mantenedores_falla(self):
        with tempfile.TemporaryDirectory() as d:
            repo = self._repo(d, maintainers=None)
            findings = self._run(repo, "slices/partials/web.md",
                                 "texto conforme\n")
            self.assertTrue(any("maintainers" in f for f in findings),
                            findings)

    def test_repo_real_propagable_limpio(self):
        catalog = Path(__file__).resolve().parents[6]
        registry = catalog / "fleet" / "registry.json"
        if not registry.is_file():
            self.skipTest("no es el arbol del catalogo team-skills")
        import glob
        paths = [Path(x) for x in glob.glob(
            str(catalog / "slices" / "partials" / "*.md"))]
        paths += [Path(x) for x in glob.glob(
            str(catalog / "slices" / "*" / "AGENTS.fragment.md"))]
        self.assertTrue(paths, "el catalogo no tiene partials ni fragmentos")
        findings = CLI._check_slice_content(catalog, paths)
        self.assertEqual(findings, [])




class GovernedMergeDocTests(unittest.TestCase):
    """#351/#354: la seccion Merge emite el camino gobernado solo cuando
    un workflow DEL CONSUMER (en su runner) invoca check_governed_merge.py;
    la deteccion es wired_files de HR-52, nunca por nombre de fichero."""

    def _repo(self, tmp, wf_name, invoke_asset):
        repo = Path(tmp)
        (repo / ".github" / "workflows").mkdir(parents=True, exist_ok=True)
        (repo / ".github" / "host-contract.json").write_text(
            json.dumps({"runner": {"label": "cadete"}}), encoding="utf-8")
        body = ("name: merge\non: workflow_dispatch\njobs:\n  merge:\n"
                "    runs-on: [self-hosted, cadete]\n    steps:\n")
        if invoke_asset:
            body += ("      - run: python3 .github/runner-controls/"
                     "check_governed_merge.py --contract @.github/host-contract.json\n")
        else:
            body += "      - run: echo sin asset\n"
        (repo / ".github" / "workflows" / wf_name).write_text(body, encoding="utf-8")
        return repo

    def test_workflow_invoking_the_asset_is_detected_regardless_of_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, "merge.yml", invoke_asset=True)
            got = CLI._governed_merge_workflow(repo)
            self.assertEqual(got, ".github/workflows/merge.yml")

    def test_workflow_named_governed_merge_without_asset_is_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, "governed-merge.yml", invoke_asset=False)
            self.assertIsNone(CLI._governed_merge_workflow(repo))

    def test_merge_lines_include_run_command_for_the_detected_workflow(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, "merge.yml", invoke_asset=True)
            text = "\n".join(CLI._operating_doc_merge_lines(repo, ["merge"]))
            self.assertIn("gh workflow run merge.yml", text)
            self.assertIn("por orden explícita y registrada de un mantenedor", text)

    def test_merge_lines_silent_without_wired_workflow(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp, "governed-merge.yml", invoke_asset=False)
            text = "\n".join(CLI._operating_doc_merge_lines(repo, ["merge"]))
            self.assertNotIn("gh workflow run", text)
            self.assertNotIn("Camino gobernado", text)


class DocsPhaseTests(unittest.TestCase):
    """#279: fase 5 — documentación operativa generada en destino + gate de
    deriva. El gate REGENERA el doc en el mismo proceso y lo compara por
    secciones; la deriva nombra la sección y la discrepancia."""

    def docs_repo(self, tmp, contract=None):
        contract = contract or complete_contract(upto=5)
        repo, _ = make_contract_repo(tmp, contract=contract)
        seed_evidence(repo, 3)
        seed_operating_doc(repo)
        return repo

    def test_docs_phase_green_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            code, out, _ = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out)

    def test_docs_phase_missing_doc_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("fase 5", err)
            self.assertIn("generate-doc", err)

    def test_docs_phase_without_declaration_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=4)
            contract["phases"]["5"] = {"something": "else"}
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("operating_doc", err)

    def test_docs_phase_unlinked_doc_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            (repo / "AGENTS.md").write_text("# consumer sin enlace\n", encoding="utf-8")
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("enlazado", err)

    def test_label_drift_names_the_section_and_value(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            drift_operating_doc(
                repo, "`bug_report.yml` → `type:bug`", "`bug_report.yml` → `type:chore`")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("deriva", err)
            self.assertIn("issues", err)
            self.assertIn("type:bug", err)

    def test_required_check_drift_names_the_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            drift_operating_doc(
                repo, "- Jobs requeridos del agregador: unit.",
                "- Jobs requeridos del agregador: unit, ghost.")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("deriva", err)
            self.assertIn("required-checks", err)

    def test_branch_regex_drift_names_the_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            doc = repo / "docs" / "ci-pattern-flow.md"
            text = doc.read_text(encoding="utf-8")
            needle = "Regex de la convención (aviso, no bloquea): `"
            k = text.index(needle) + len(needle)
            end = text.index("`", k)
            drifted = text[:k] + "^main$" + text[end:]
            doc.write_text(drifted, encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("deriva", err)
            self.assertIn("branch", err)

    def test_unreadable_params_input_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            (repo / "ci-pattern.yaml").write_text("P09_labels_type:\n\tbad\n",
                                                  encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 2, err)

    def test_missing_params_input_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            (repo / "ci-pattern.yaml").unlink()
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 2, err)

    def test_unreadable_host_contract_is_exit_2(self):
        # a nivel gate (--phase 5) la fase 4 corta antes con el JSON roto;
        # la misma entrada ilegible es exit 2 en el generador.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            (repo / ".github" / "host-contract.json").write_text("{no json", encoding="utf-8")
            code, _, err = run_cli(
                ["adoption", "generate-doc", str(repo),
                 "--out", str(repo / "docs" / "ci-pattern-flow.md")])
            self.assertEqual(code, 2, err)


class DocsLocationTests(unittest.TestCase):
    """#319 (HR-48): la documentación de un proyecto vive en SU repositorio.
    Ni `AGENTS.md`, ni el documento operativo, ni los ficheros que estos
    enlazan pueden mandar documentación a una ruta absoluta o externa al
    repo: la fase 5 lo comprueba y la fase 3 sobre `AGENTS.md`."""

    EXTERNAL = ("New or updated project documentation belongs in "
                "`C:\\00repos\\documentacion\\OPENSPEC\\00_CADETE`, not in this repo.\n")

    def docs_repo(self, tmp, agents_extra=""):
        contract = complete_contract(upto=5)
        repo, _ = make_contract_repo(tmp, contract=contract)
        seed_evidence(repo, 3)
        seed_operating_doc(repo)
        if agents_extra:
            agents = repo / "AGENTS.md"
            agents.write_text(agents.read_text(encoding="utf-8") + agents_extra,
                              encoding="utf-8")
            seed_manifest(repo)  # AGENTS.md es artefacto sellado (HR-46)
        return repo

    def test_external_docs_path_in_agents_md_names_file_and_line(self):
        # la fase 3 es la que gobierna AGENTS.md, así que frena antes que la 5
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp, self.EXTERNAL)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("fase 3", err)
            self.assertRegex(err, r"AGENTS\.md:\d+")
            self.assertIn("00_CADETE", err)
            self.assertIn("HR-48", err)

    def test_phase_five_relative_docs_paths_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(
                tmp, "Toda la documentación del proyecto vive en este repositorio, "
                     "bajo `docs/`; los post-mortems en "
                     "`docs/postmortems/2026-10-06-incidente.md`.\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_three_external_docs_path_in_agents_md_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            agents = repo / "AGENTS.md"
            agents.write_text(agents.read_text(encoding="utf-8") + self.EXTERNAL,
                              encoding="utf-8")
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("fase 3", err)
            self.assertRegex(err, r"AGENTS\.md:\d+")
            self.assertIn("HR-48", err)

    def test_phase_five_linked_file_with_external_docs_path_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            notes = repo / "docs" / "notas.md"
            notes.write_text("La documentación se guarda en "
                             "`\\\\datoste\\documentacion` según la convención.\n",
                             encoding="utf-8")
            agents = repo / "AGENTS.md"
            agents.write_text(agents.read_text(encoding="utf-8")
                              + "Notas del equipo: [notas](docs/notas.md)\n",
                              encoding="utf-8")
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("fase 5", err)
            self.assertRegex(err, r"docs/notas\.md:\d+")
            self.assertIn("HR-48", err)

    def test_generated_doc_states_where_documentation_lives(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.docs_repo(tmp)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            self.assertIn("<!-- ci-pattern-doc:begin docs-location -->", doc)
            self.assertIn("docs/postmortems/", doc)
            agents = (repo / "AGENTS.md").read_text(encoding="utf-8")
            self.assertIn("Dónde vive la documentación", agents)


class DocFidelityTests(unittest.TestCase):
    """#356: el documento operativo generado solo puede nombrar comandos que existan
    en el consumer (o que se instalen en él) y los tipos permitidos salen del
    contrato del host, no del mapeo de formularios."""

    def consumer_with_contract_types(self, tmp):
        """El contrato del host declara cuatro etiquetas `type:*` y el formulario
        no declara ninguna (`labels: []`)."""
        contract = complete_contract(upto=5)
        repo, _ = make_contract_repo(tmp, contract=contract)
        host_path = repo / ".github" / "host-contract.json"
        host = json.loads(host_path.read_text(encoding="utf-8"))
        host["labels"] = [{"name": name, "class": "host-enforced"}
                          for name in ("type:bug", "type:chore", "type:docs", "type:feature")]
        host_path.write_text(json.dumps(host), encoding="utf-8")
        (repo / "ci-pattern.yaml").write_text(
            "P09_labels_type:\n  issue-canonical.yml: type:bug\n", encoding="utf-8")
        templates = repo / ".github" / "ISSUE_TEMPLATE"
        templates.mkdir(parents=True, exist_ok=True)
        (templates / "issue-canonical.yml").write_text(
            "name: Issue canónica\nlabels: []\n", encoding="utf-8")
        seed_evidence(repo, 3)
        seed_operating_doc(repo)
        return repo

    def test_doc_takes_the_type_labels_from_the_host_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.consumer_with_contract_types(tmp)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            line = next(item for item in doc.splitlines() if "Etiquetas `type:*` orientativas" in item)
            for label in ("type:bug", "type:chore", "type:docs", "type:feature"):
                self.assertIn(label, line)

    def test_doc_falls_back_to_the_form_mapping_without_contract_types(self):
        """Sin etiquetas de tipo en el contrato del host, el doc sigue listando el
        mapeo de formularios declarado (respaldo de #356)."""
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            line = next(item for item in doc.splitlines() if "Etiquetas `type:*` orientativas" in item)
            self.assertIn("type:bug", line)
            self.assertIn("type:feature", line)

    def governed_merge_repo(self, tmp):
        """Consumer con camino gobernado: runner declarado y workflow que invoca
        el merge gobernado en SU runner."""
        repo = self.consumer_with_contract_types(tmp)
        host = json.loads((repo / ".github" / "host-contract.json").read_text(encoding="utf-8"))
        host["runner"] = {"label": "fixture"}
        (repo / ".github" / "host-contract.json").write_text(json.dumps(host), encoding="utf-8")
        controls = repo / ".github" / "runner-controls"
        controls.mkdir(parents=True, exist_ok=True)
        (controls / "check_governed_merge.py").write_text("# gate\n", encoding="utf-8")
        (repo / ".github" / "workflows" / "governed-merge.yml").write_text(
            "name: Governed merge\non:\n  workflow_dispatch:\n    inputs:\n      pr:\n"
            "        required: true\n        type: string\njobs:\n  governed-merge:\n"
            "    runs-on: [self-hosted, fixture]\n    steps:\n      - name: Merge\n"
            "        run: |\n          python3 .github/runner-controls/check_governed_merge.py \\\n"
            "            --repo x/y --pr 1 --contract .github/host-contract.json\n",
            encoding="utf-8")
        seed_operating_doc(repo)
        return repo

    def test_merge_rule_names_the_role_not_a_person(self):
        """#356 (decisión del usuario, 2026-10-07): el merge gobernado lo lanza un
        mantenedor del repositorio, no «el operador»."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.governed_merge_repo(tmp)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            line = next(item for item in doc.splitlines() if "Camino gobernado" in item)
            self.assertIn("mantenedor", line)
            self.assertNotIn("operador", line)

    def test_doc_names_no_command_that_the_consumer_does_not_have(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.consumer_with_contract_types(tmp)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            self.assertNotIn("assets/branch-name.sh", doc)
            self.assertIn("git switch -c", doc)


class IncidentRecordTests(unittest.TestCase):
    """#322 (HR-21 ampliada): de un incidente se guardan cuatro niveles —el
    post-mortem (parámetro 9), el registro operativo (por defecto
    `docs/incidents/` del propio repo), los secretos (gestor de secretos) y
    los datos personales (sistema de tickets)—. El gate bloquea SOLO
    secretos y datos personales, y la evidencia declarada sin hash ni
    ubicación: las IPs, los hostnames y los comandos internos están
    permitidos."""

    def repo_with_incidents(self, tmp, postmortem=None, record=None,
                            params_extra=""):
        contract = complete_contract(upto=5)
        repo, _ = make_contract_repo(tmp, contract=contract)
        seed_evidence(repo, 3)
        seed_operating_doc(repo)
        if params_extra:
            params = repo / "ci-pattern.yaml"
            params.write_text(params.read_text(encoding="utf-8") + params_extra,
                              encoding="utf-8")
            seed_manifest(repo)  # ci-pattern.yaml es artefacto sellado
        if postmortem:
            touch(repo, "docs/postmortems/2026-10-06-incidente.md", postmortem)
        if record:
            touch(repo, "docs/incidents/2026-10-06-incidente.md", record)
        return repo

    def repo_with_params(self, tmp, params_extra):
        # #915: el documento operativo se genera DESPUÉS de escribir el
        # contrato, para que lo que se prueba sea el contrato declarado.
        contract = complete_contract(upto=5)
        repo, _ = make_contract_repo(tmp, contract=contract)
        seed_evidence(repo, 3)
        params = repo / "ci-pattern.yaml"
        params.write_text(params.read_text(encoding="utf-8") + params_extra,
                          encoding="utf-8")
        seed_operating_doc(repo)
        seed_manifest(repo)
        return repo

    def test_secret_in_postmortem_names_file_and_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, postmortem="# Incidente\n\n- password: hunter2\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/postmortems/2026-10-06-incidente\.md:3")
            self.assertIn("HR-21", err)

    def test_prose_with_clave_colon_is_not_a_credential(self):
        # #350: el habla común usa "clave:" para "punto clave". Sin valor con
        # estructura de secreto no es un hallazgo; antes el gate obligaba a
        # reescribir prosa ajena para pasar.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, postmortem="# Incidente\n\n- Pendiente clave: backups "
                                "automatizados en produccion (#280)\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, err)

    def test_clave_with_secret_shaped_value_is_a_finding(self):
        # #350: el arreglo no abre el agujero. Una credencial colgada de
        # `clave:` se sigue señalando con fichero y línea.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, postmortem="# Incidente\n\n- clave: Abc12345xyz\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/postmortems/2026-10-06-incidente\.md:3")

    def test_clave_with_plain_word_is_not_a_credential(self):
        # #350: la frontera exacta del patrón nuevo — una palabra suelta, sin
        # dígito ni comillas, no tiene estructura de secreto.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, postmortem="# Incidente\n\n- clave: automatizaciones\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, err)

    def test_personal_email_in_operational_record_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, record="# Registro\n\n- Contacto: juan.perez@gmail.com\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/incidents/2026-10-06-incidente\.md:3")

    def test_evidence_without_hash_or_location_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp, record="# Registro\n\n- Evidencia: el log del pod, adjunto.\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("evidencia", err.lower())

    def test_operational_detail_with_ips_and_commands_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp,
                postmortem=("# Incidente\n\n- ID: INC000017215662; ticket BMC "
                            "TKT-9931; SHA 4f2b1c9\n"
                            "- Evidencia: rca.zip — sha256 " + "a" * 64 + "\n"),
                record=("# Registro operativo\n\n- Nodo 10.20.30.41 (cadete-prod-01), "
                        "PVC cadete-mysql\n"
                        "- Comando: oc exec cadete-mysql-0 -- mysqladmin status\n"
                        "- Ticket BMC TKT-9931: esperando respuesta de IT\n"))
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_configured_incident_dir_is_scanned(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp,
                params_extra="P49_incident_log_dir: docs/bitacora\n")
            touch(repo, "docs/bitacora/2026-10-06-otro.md",
                  "# Bitacora\n\n- api_key: sk-live-1234567890\n")  # gitleaks:allow — fixture sintetico: el valor es un marcador del test, no un secreto real
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/bitacora/2026-10-06-otro\.md:3")

    def test_postmortem_and_record_sharing_a_path_fails(self):
        # #927 (HR-21): el post-mortem y el registro operativo son niveles
        # distintos; compartir ruta es un hallazgo de fase 5 (exit 1) que
        # nombra el parámetro a cambiar. Hasta #915 era solo un aviso.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_params(
                tmp, "P49_incident_log_dir: docs/postmortems\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, out + err)
            self.assertIn("hallazgo: fase 5", err)
            self.assertIn("P49_incident_log_dir", err)
            self.assertIn("docs/postmortems", err)

    def test_postmortem_collision_appears_alongside_other_phase_5_findings(self):
        # #927 (HR-21): la colisión es un hallazgo de fase 5 más, no un
        # `return` aparte. Cuando hay otro hallazgo de fase 5 (un secreto
        # embebido en el post-mortem, que también lo es) y la colisión, la
        # salida lista los dos en una sola pasada, con el prefijo `hallazgo:`
        # y sin un `return` posterior que oculte la colisión. El valor del
        # secreto NO se imprime (HR-21: las credenciales no salen a stderr);
        # el gate nombra el fichero y la línea.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(
                tmp,
                postmortem="# Incidente\n\n- password: hunter2\n",
                params_extra="P49_incident_log_dir: docs/postmortems\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, out + err)
            # Ambos hallazgos en la misma pasada:
            self.assertIn("P49_incident_log_dir", err)
            self.assertIn("docs/postmortems", err)
            # El log del secreto nombra el fichero y la línea, no el valor;
            # la colisión nombra el parámetro a cambiar. Los dos llevan
            # prefijo hallazgo:
            self.assertRegex(err, r"hallazgo: fase 5:.*P49_incident_log_dir")
            self.assertRegex(err, r"hallazgo: fase 5:.*docs/postmortems/[^ ]*\.md:3")

    def test_distinct_record_path_has_no_collision_warning(self):
        # #915: con rutas distintas no hay aviso de colisión.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_params(
                tmp, "P49_incident_log_dir: docs/incidents\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)
            self.assertNotIn("aviso", err)

    def test_local_unversioned_record_is_described_and_not_required(self):
        # #915: un registro operativo local no tiene ruta en el repo: no se
        # exige ni se escanea, y el documento lo describe con la frase, no con
        # una ruta.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_params(
                tmp, "P55_incident_log_versioning: local-unversioned\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            self.assertIn("local y no versionado", doc)
            self.assertNotIn("docs/incidents/", doc)

    def test_local_unversioned_record_is_not_scanned(self):
        # #915: lo que el consumer declara como local no vive en el repo: el
        # gate no lo escanea, así que un fichero de esa ruta no es hallazgo.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_params(
                tmp, "P55_incident_log_versioning: local-unversioned\n")
            touch(repo, "docs/incidents/run.log", "password: hunter2\n")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_params_validate_accepts_local_unversioned_and_rejects_other_values(self):
        # #915: el vocabulario de P55 es cerrado: versioned o local-unversioned.
        with tempfile.TemporaryDirectory() as tmp:
            ok = Path(tmp) / "ok.yaml"
            ok.write_text("P55_incident_log_versioning: local-unversioned\n",
                          encoding="utf-8")
            bad = Path(tmp) / "bad.yaml"
            bad.write_text("P55_incident_log_versioning: quizas\n", encoding="utf-8")
            # El fichero es parcial: `validate` exige los obligatorios, así que
            # el criterio es que P55 no aparezca entre los errores.
            code, out, err = run_cli(["params", "validate", str(ok)])
            self.assertNotIn("P55_incident_log_versioning", out + err)
            code, out, err = run_cli(["params", "validate", str(bad)])
            self.assertEqual(code, 1, out + err)
            self.assertRegex(out + err, r"P55_incident_log_versioning: debe ser uno de")

    def test_scan_limit_exceeded_is_not_a_silent_pass(self):
        # 201 ficheros: el límite no puede dar verde parcial saltándose el
        # resto en silencio.
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(tmp)
            for i in range(201):
                touch(repo, f"docs/incidents/2026-10-06-{i:03d}.md", f"# {i}\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 2, err)
            self.assertIn("201", err)
            self.assertIn("l\u00edmite", err.lower())

    def test_secret_in_log_file_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(tmp)
            touch(repo, "docs/incidents/run.log",
                  "2026-10-06T10:00 curl -H 'Authorization: Bearer "
                  "ghp_AbCdEfGhIjKlMnOpQrStUvWxYz0123456789'\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/incidents/run\.log:1")

    def test_secret_in_json_and_yaml_is_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(tmp)
            touch(repo, "docs/incidents/estado.json", '{"access_token": "AKIAIOSFODNN7EXAMPLE"}\n')
            touch(repo, "docs/incidents/estado.yaml", "claves:\n  password: hunter2\n")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertRegex(err, r"docs/incidents/estado\.(json|yaml):\d")

    def test_binary_file_is_skipped(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(tmp)
            (repo / "docs" / "incidents").mkdir(parents=True, exist_ok=True)
            (repo / "docs" / "incidents" / "volcado.log").write_bytes(b"\x7fELF\x00password: hunter2\x00")
            code, out, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_generated_doc_has_the_incident_recipe(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.repo_with_incidents(tmp)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            self.assertIn("<!-- ci-pattern-doc:begin incidents -->", doc)
            self.assertIn("docs/incidents/", doc)
            self.assertIn("docs/postmortems/", doc)
            self.assertIn("gestor de secretos", doc)


PLAN_403 = {
    "message": "Upgrade to GitHub Pro or make this repository public to enable this feature.",
    "documentation_url":
        "https://docs.github.com/rest/branches/branch-protection#get-branch-protection",
    "status": "403",
}


class PlanUnavailableSnapshotTests(unittest.TestCase):
    """#330: en un repo privado sin plan, el GET de protección de rama y el de
    rulesets devuelven el 403 de plan. Ese cuerpo es EVIDENCIA de capacidad no
    disponible y vale como instantánea **solo** si el contrato del host declara
    esa área `unavailable`; con `available` es drift (exit 1) y sin declaración
    es duda (exit 2)."""

    def plan_repo(self, tmp, capabilities="unset"):
        contract = complete_contract(upto=3)
        repo, _ = make_contract_repo(tmp, contract=contract)
        if capabilities != "unset":
            host_path = repo / ".github" / "host-contract.json"
            host = json.loads(host_path.read_text(encoding="utf-8"))
            host["host_capabilities"] = capabilities
            touch(repo, ".github/host-contract.json",
                  json.dumps(host, indent=2, ensure_ascii=False))
        # evidencia 1..3 + instantáneas base + manifiesto: el sello incluye el
        # contrato del host ya editado, así que se siembra DESPUÉS de editarlo
        seed_evidence(repo, 3)
        # el host responde el 403 de plan en las DOS áreas (no van selladas:
        # las instantáneas del host no forman parte del manifiesto)
        touch(repo, "docs/adoption/branch-protection.json",
              json.dumps(PLAN_403, indent=2, ensure_ascii=False))
        touch(repo, "docs/adoption/rulesets.json",
              json.dumps(PLAN_403, indent=2, ensure_ascii=False))
        return repo

    def test_plan_403_with_unavailable_capability_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.plan_repo(tmp, {"branch_protection": "unavailable",
                                       "rulesets": "unavailable",
                                       "required_checks": "unavailable"})
            code, out, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_plan_403_with_available_capability_is_drift_naming_the_area(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.plan_repo(tmp, {"branch_protection": "available",
                                       "rulesets": "available"})
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("branch_protection", err)
            self.assertIn("403", err)

    def test_plan_403_without_capability_declaration_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.plan_repo(tmp)  # sin host_capabilities
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2, err)
            self.assertIn("host_capabilities", err)

    def test_plan_403_with_unknown_capability_is_exit_2_naming_the_value(self):
        # un valor fuera del vocabulario (`available`/`unavailable`) no se puede
        # interpretar: es duda, no drift
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.plan_repo(tmp, {"branch_protection": "foo",
                                       "rulesets": "foo"})
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2, err)
            self.assertIn("foo", err)
            self.assertIn("host_capabilities", err)

    def test_real_rulesets_payload_still_requires_a_list(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.plan_repo(tmp, {"branch_protection": "unavailable",
                                       "rulesets": "unavailable"})
            touch(repo, "docs/adoption/rulesets.json", '{"name": "x"}\n')
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 2, err)
            self.assertIn("array", err)


class CliSurfaceTests(unittest.TestCase):
    def test_help_is_accepted(self):
        # auditoría de contenido: `adoption check --help` es documentación,
        # no un error de uso.
        for argv in (["adoption", "-h"], ["adoption", "check", "--help"]):
            code, out, _ = run_cli(argv)
            self.assertEqual(code, 0, out)
            self.assertIn("adoption check", out)

    def test_phase_must_be_number(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
            code, _, _ = run_cli(["adoption", "check", "--phase", "cero", str(repo)])
            self.assertEqual(code, 2)

    def test_phase_out_of_range_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
            code, _, _ = run_cli(["adoption", "check", "--phase", "8", str(repo)])
            self.assertEqual(code, 2)

    def test_missing_repo_argument_is_exit_2(self):
        code, _, _ = run_cli(["adoption", "check", "--all"])
        self.assertEqual(code, 2)

    def test_unknown_adoption_subcommand_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
            code, _, _ = run_cli(["adoption", "mutate", str(repo)])
            self.assertEqual(code, 2)


def make_docs_repo(tmp):
    """Repo falso con las entradas del generador del documento operativo
    (#279): ci-pattern.yaml en la raíz, contrato de host, política de jobs
    requeridos y formularios de issue."""
    repo = Path(tmp) / "consumer"
    (repo / ".github" / "ISSUE_TEMPLATE").mkdir(parents=True)
    (repo / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml").write_text(
        "name: Bug report\n", encoding="utf-8")
    (repo / ".github" / "ISSUE_TEMPLATE" / "feature_request.yml").write_text(
        "name: Feature request\n", encoding="utf-8")
    (repo / "ci-pattern.yaml").write_text(
        "P09_labels_type:\n"
        "  bug_report.yml: type:bug\n"
        "  feature_request.yml: type:feature\n",
        encoding="utf-8")
    (repo / ".github" / "host-contract.json").write_text(json.dumps({
        "contract_version": 1,
        "labels": [
            {"name": "chain:partial", "class": "host-enforced"},
            {"name": "status:approved", "class": "host-enforced"},
            {"name": "type:bug", "class": "host-enforced"},
            {"name": "type:feature", "class": "host-enforced"},
        ],
        "required_checks": [],
        "merge_methods": {
            "allow_merge_commit": {"declared": True, "class": "host-enforced"},
            "allow_squash_merge": {"declared": False, "class": "host-enforced"},
            "allow_rebase_merge": {"declared": False, "class": "host-enforced"},
            "allow_auto_merge": {"declared": True, "class": "host-enforced"},
        },
        "protection": {},
        "rulesets": [],
    }), encoding="utf-8")
    (repo / ".github" / "required-jobs-policy.json").write_text(json.dumps({
        "required_jobs": ["unit", "lint"],
        "events": {e: {"accepted_skips": []} for e in
                   ("pull_request", "push", "schedule", "workflow_dispatch")},
        "local": {"entrypoint": "testing/run-unit-suites.sh",
                  "ci": {"workflow": ".github/workflows/ci.yml", "job": "unit"}},
        "exclusions": [],
    }), encoding="utf-8")
    (repo / "AGENTS.md").write_text("# consumer\n", encoding="utf-8")
    return repo


class OperatingDocTests(unittest.TestCase):
    """#279: generador determinista del documento operativo del consumer."""

    DOC_SECTIONS = (
        "issues", "branch", "commits", "pull-request",
        "required-checks", "merge", "release-hotfix", "friction",
        "lifecycle",
    )

    def test_generate_doc_help_is_accepted(self):
        # nit del auditor: `-h`/`--help` es documentación, no un error de
        # uso; misma regla que `adoption check`.
        for argv in (["adoption", "generate-doc", "-h"],
                     ["adoption", "generate-doc", "--help"]):
            code, out, _ = run_cli(argv)
            self.assertEqual(code, 0, out)
            self.assertIn("generate-doc", out)

    def test_generate_doc_is_deterministic(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            out1 = Path(tmp) / "doc1.md"
            out2 = Path(tmp) / "doc2.md"
            code1, _, err1 = run_cli(
                ["adoption", "generate-doc", str(repo), "--out", str(out1)])
            code2, _, err2 = run_cli(
                ["adoption", "generate-doc", str(repo), "--out", str(out2)])
            self.assertEqual(code1, 0, err1)
            self.assertEqual(code2, 0, err2)
            self.assertEqual(
                out1.read_bytes(), out2.read_bytes(),
                "el documento operativo no es determinista byte a byte")

    def test_generate_doc_prose_has_no_decorative_emoji_or_shouting_caps(self):
        # #652: check_alantyle (ALAN002/ALAN003) rechaza el documento generado.
        # El 📍 de Chain Context vive en el diagrama de la plantilla de PR, no
        # en la prosa, y las instrucciones van en minúsculas.
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            out = Path(tmp) / "docs" / "ci-pattern-flow.md"
            code, _, err = run_cli(
                ["adoption", "generate-doc", str(repo), "--out", str(out)])
            self.assertEqual(code, 0, err)
            text = out.read_text(encoding="utf-8")
            self.assertNotIn("\U0001F4CD", text,
                             "el 📍 no puede ir en la prosa del documento generado")
            shouting = re.findall(
                r"\b(GENERADO|CERO|CONTENIDO|NUNCA|ANTES|SIEMPRE|ESTE)\b", text)
            self.assertEqual(shouting, [],
                             "mayúsculas de prosa en el documento generado")

    def test_generate_doc_out_writes_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            out = Path(tmp) / "docs" / "ci-pattern-flow.md"
            code, _, err = run_cli(
                ["adoption", "generate-doc", str(repo), "--out", str(out)])
            self.assertEqual(code, 0, err)
            text = out.read_text(encoding="utf-8")
            for section in self.DOC_SECTIONS:
                self.assertIn(
                    f"ci-pattern-doc:begin {section}", text,
                    f"falta la sección '{section}' del documento operativo")
                self.assertIn(f"ci-pattern-doc:end {section}", text)
            # Datos reales del consumer, no prosa genérica: regex de rama
            # (default del esquema P01), jobs requeridos y mapeo formulario
            # → etiqueta.
            self.assertIn("(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)", text)
            self.assertIn("unit", text)
            self.assertIn("lint", text)
            self.assertIn("type:bug", text)
            self.assertIn("type:feature", text)

    def test_generate_doc_without_out_is_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            code, _, err = run_cli(["adoption", "generate-doc", str(repo)])
            self.assertEqual(code, 2, err)
            self.assertIn("--out", err)


def _lifecycle_section(text):
    """Cuerpo de la sección 'lifecycle' del documento generado."""
    start = text.index("<!-- ci-pattern-doc:begin lifecycle -->")
    end = text.index("<!-- ci-pattern-doc:end lifecycle -->")
    return text[start:end]


def _pull_request_section(text):
    """Cuerpo de la sección 'pull-request' del documento generado."""
    start = text.index("<!-- ci-pattern-doc:begin pull-request -->")
    end = text.index("<!-- ci-pattern-doc:end pull-request -->")
    return text[start:end]


class OperatingDocAuditHolesTests(unittest.TestCase):
    """Oleada 1 (#873): cubre los cuatro huecos estructurales que la
    auditoría del repo cadete encontró en el documento operativo que el
    patrón genera. RED primero sobre el `bin/ci-pattern` actual; GREEN
    cuando se actualiza el generador.

    Fuera de alcance (a oleada 2): la condicionalidad de la línea HR-44
    sobre el cableado del gate `changelog-fragment` requiere detectar el
    cableado desde los datos que ya existen (workflows del consumer) y
    mover el esquema del host-contract queda descartado por decisión del
    owner."""

    def _generate(self, repo):
        out = repo.parent / "doc.md"
        code, _, err = run_cli(
            ["adoption", "generate-doc", str(repo), "--out", str(out)])
        self.assertEqual(code, 0, err)
        return out.read_text(encoding="utf-8")

    def test_gh_issue_create_carries_type_label_marker(self):
        """S1: etapa 1 lleva `gh issue create ... --label type:<tipo>` como
        MARCADOR LITERAL cuando el host-contract declara al menos una
        etiqueta `type:*`. La receta NO concreta la etiqueta: el agente
        elige entre los valores de la sección 4, no toma la primera del
        contrato (defecto del round 1)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            text = self._generate(repo)
            lifecycle = _lifecycle_section(text)
            self.assertIn(
                "gh issue create --title ", lifecycle)
            self.assertIn(
                "--label type:<tipo>", lifecycle,
                "la etapa 1 del doc no lleva el marcador literal "
                "`--label type:<tipo>` (S1)")
            for forbidden in ("--label type:bug", "--label type:feature"):
                self.assertNotIn(
                    forbidden, lifecycle,
                    f"la etapa 1 del doc concreta una etiqueta ({forbidden!r}); "
                    "un agente que copiara la receta la aplicaría a todo: "
                    "use el marcador literal `--label type:<tipo>`")

    def test_gh_pr_create_carries_type_label_marker(self):
        """S2: etapa 6 lleva `gh pr create ... --label type:<tipo>` como
        MARCADOR LITERAL. Misma regla que S1: NUNCA concreta la etiqueta."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            text = self._generate(repo)
            lifecycle = _lifecycle_section(text)
            self.assertIn("gh pr create", lifecycle)
            self.assertIn(
                "--label type:<tipo>", lifecycle,
                "la etapa 6 del doc no lleva el marcador literal "
                "`--label type:<tipo>` (S2)")
            for forbidden in ("--label type:bug", "--label type:feature"):
                self.assertNotIn(
                    forbidden, lifecycle,
                    f"la etapa 6 del doc concreta una etiqueta ({forbidden!r}); "
                    "use el marcador literal `--label type:<tipo>`")

    def test_no_type_label_when_host_contract_omits_type(self):
        """S1/S2 (inverso): sin etiquetas `type:*` en el host-contract, los
        comandos `gh issue create` y `gh pr create` se emiten SIN la bandera
        `--label` (no se inventa)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            # quita todas las etiquetas `type:*` del host-contract; deja
            # el resto (status:approved, chain:partial).
            host_path = repo / ".github" / "host-contract.json"
            host = json.loads(host_path.read_text(encoding="utf-8"))
            host["labels"] = [label for label in (host.get("labels") or [])
                              if not (isinstance(label, dict)
                                      and isinstance(label.get("name"), str)
                                      and label["name"].startswith("type:"))]
            host_path.write_text(json.dumps(host), encoding="utf-8")
            text = self._generate(repo)
            lifecycle = _lifecycle_section(text)
            self.assertIn("gh issue create", lifecycle)
            self.assertNotIn(
                "--label type:<tipo>", lifecycle,
                "sin `type:*` en el host-contract, el `gh issue create` del "
                "doc no debe llevar `--label type:<tipo>` (S1)")
            self.assertIn("gh pr create", lifecycle)
            self.assertNotIn(
                "--label type:<tipo>", lifecycle,
                "sin `type:*` en el host-contract, el `gh pr create` del "
                "doc no debe llevar `--label type:<tipo>` (S2)")

    def test_command_lines_keep_backticks_balanced(self):
        """S1/S2: cada línea de `gh issue create` / `gh pr create` tiene un
        número par de backticks. Un backtick suelto abre un bloque de código
        que se come el resto de la sección del doc."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            lifecycle = _lifecycle_section(self._generate(repo))
            checked = 0
            for line in lifecycle.splitlines():
                if "gh issue create" in line or "gh pr create" in line:
                    checked += 1
                    self.assertEqual(line.count("`") % 2, 0,
                                     f"backticks desbalanceados: {line!r}")
            self.assertEqual(checked, 2)

    def test_worktree_path_uses_worktrees_sibling(self):
        """S3 (etapa 3): la ruta del worktree va a `../<repo>-worktrees/<número>-<tema>`,
        como el hook pre-commit de agent-guards y worktree-reorg-per-project v2.0.
        Antes decía `../<repo>-<número>-<tema>` (layout huérfano)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            text = self._generate(repo)
            lifecycle = _lifecycle_section(text)
            self.assertIn(
                "../<repo>-worktrees/", lifecycle,
                "la etapa 3 no usa el layout `worktrees/` sibling-container; "
                "la receta contradice worktree-reorg-per-project v2.0 (S3)")
            self.assertNotIn(
                "../<repo>-<número>-<tema>`", lifecycle,
                "la etapa 3 sigue usando el layout antiguo `../<repo>-<número>-<tema>` "
                "que worktree-reorg-per-project v2.0 retiró")

    def test_worktree_remove_uses_worktrees_sibling(self):
        """S3 (etapa 10): limpieza usa la misma ruta `../<repo>-worktrees/<número>-<tema>`."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            text = self._generate(repo)
            lifecycle = _lifecycle_section(text)
            occurrences = lifecycle.count("../<repo>-worktrees/")
            self.assertGreaterEqual(
                occurrences, 2,
                "la sección lifecycle debe llevar la ruta `worktrees/` en "
                "etapas 3 y 10 (S3)")
            self.assertNotIn(
                "../<repo>-<número>-<tema>`", lifecycle,
                "la sección lifecycle aún contiene el layout antiguo (S3)")

    def test_hr_gate_matrix_line_withdrawn(self):
        """S4: la sección 'pull-request' del doc NO contiene la línea de la
        matriz HR→gate, retirada del catálogo en #513/#540/#542. La receta
        apuntaba a `references/hr-gate-matrix.json`, que ya no existe."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            text = self._generate(repo)
            pull_request = _pull_request_section(text)
            self.assertNotIn(
                "matriz HR", pull_request,
                "la sección pull-request sigue mencionando la matriz HR→gate (S4); "
                "el fichero fue retirado del catálogo en #513/#540/#542")
            self.assertNotIn(
                "hr-gate-matrix.json", pull_request,
                "la sección pull-request apunta a un fichero retirado (S4)")


class LifecycleRecipeTests(unittest.TestCase):
    """#305: la sección 'lifecycle' del documento operativo recorre el ciclo
    completo (issue → aprobación → worktree → commits → preflight local →
    PR → checks → CI verde → merge → limpieza) con comando exacto y
    condición de «hecho» por etapa, generada desde los datos del consumer.
    Sin pasos implícitos: una receta ausente o vacía es hallazgo del gate."""

    def doc_of(self, tmp, extra_jobs=()):
        repo = make_docs_repo(tmp)
        if extra_jobs:
            policy_path = repo / ".github" / "required-jobs-policy.json"
            policy = json.loads(policy_path.read_text(encoding="utf-8"))
            policy["required_jobs"] += list(extra_jobs)
            policy_path.write_text(json.dumps(policy), encoding="utf-8")
        out = Path(tmp) / "doc.md"
        code, _, err = run_cli(
            ["adoption", "generate-doc", str(repo), "--out", str(out)])
        self.assertEqual(code, 0, err)
        text = out.read_text(encoding="utf-8")
        start = text.index("<!-- ci-pattern-doc:begin lifecycle -->")
        end = text.index("<!-- ci-pattern-doc:end lifecycle -->")
        return text[start:end]

    def test_ten_stages_in_order(self):
        body = self.doc_of(tempfile.mkdtemp())
        stages = ("Issue", "Aprobación", "Worktree", "Commits",
                  "Integración de la base y preflight",
                  "Pull request", "Checks", "CI en verde", "Merge",
                  "Limpieza")
        positions = [body.find(f"**{n}. {name}")
                     for n, name in enumerate(stages, start=1)]
        for pos, (n, name) in zip(positions, enumerate(stages, start=1), strict=True):
            self.assertGreaterEqual(pos, 0, f"falta la etapa {n} ({name})")
        self.assertEqual(positions, sorted(positions),
                         "las etapas del ciclo no están en orden")

    def test_every_stage_has_done_condition(self):
        body = self.doc_of(tempfile.mkdtemp())
        self.assertEqual(body.count("Hecho:"), 10,
                         "cada etapa exige su condición de «hecho»")

    def test_approval_stage_is_a_convention_not_a_gate(self):
        # R1 (#1093): `status:approved` la puede poner cualquiera, agentes
        # incluidos, y ningún gate la exige. La etapa 2 no puede mandar a
        # parar ni a negar la rama o el PR por la etiqueta.
        body = self.doc_of(tempfile.mkdtemp())
        stage = body[body.index("**2. Aprobación"):body.index("**3. Worktree.**")]
        for fragment in (
            "la puede poner cualquiera, agentes incluidos",
            "ningún gate la exige",
            "el enlace a la issue, el tamaño y los commits convencionales",
            "Hecho: la etiqueta `status:approved` es opcional",
        ):
            self.assertIn(fragment, stage, f"la etapa 2 no dice: {fragment}")
        for retired in (
            "Nunca se auto-aprueba",
            "no cree rama ni PR",
            "pare y pregunte",
            "la aprueba un mantenedor",
        ):
            self.assertNotIn(retired, stage, f"la etapa 2 aún dice: {retired}")

    def test_commands_are_literal_and_complete(self):
        body = self.doc_of(tempfile.mkdtemp())
        for fragment in (
            "gh issue list --search", "--state all", "--limit 20",
            "gh issue create", "gh issue view",
            "git fetch origin", "git worktree add", "git worktree list",
            "git log --oneline",
            "bash testing/run-unit-suites.sh",
            "gh pr create", "gh pr checks", "gh pr view",
            "git worktree remove",
        ):
            self.assertIn(fragment, body, f"falta el comando literal: {fragment}")

    def test_uses_consumer_data_not_generic_prose(self):
        body = self.doc_of(tempfile.mkdtemp())
        # regex de rama (default P01), presupuesto (P06), etiqueta de
        # aprobación (P10) y método de merge del contrato del host.
        self.assertIn("(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)", body)
        self.assertIn("400 líneas", body)
        self.assertIn("status:approved", body)
        self.assertIn("merge commit", body)

    def test_missing_local_entrypoint_is_exit_2(self):
        # sin entrada `local` en la política de jobs, la receta de
        # preflight no puede ser literal: el generador falla cerrado.
        with tempfile.TemporaryDirectory() as tmp:
            repo = make_docs_repo(tmp)
            policy = json.loads(
                (repo / ".github" / "required-jobs-policy.json")
                .read_text(encoding="utf-8"))
            del policy["local"]
            (repo / ".github" / "required-jobs-policy.json").write_text(
                json.dumps(policy), encoding="utf-8")
            code, _, err = run_cli(
                ["adoption", "generate-doc", str(repo),
                 "--out", str(Path(tmp) / "doc.md")])
            self.assertEqual(code, 2, err)
            self.assertIn("entrypoint", err)

    def test_deleted_recipe_line_is_gate_finding(self):
        # fixture RED del criterio: una receta eliminada del documento
        # generado es deriva nombrada por el gate de la fase 5.
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, cp = make_contract_repo(tmp, contract=contract)
            init_git_with_commit(repo)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            doc_rel = contract["phases"]["5"]["operating_doc"]["doc"]
            doc = repo / doc_rel
            text = doc.read_text(encoding="utf-8")
            text = text.replace("git worktree add", "git worktree-REDACTED")
            doc.write_text(text, encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("lifecycle", err)

class BranchNameShapeTests(unittest.TestCase):
    """#842: el apartado de rama del documento operativo describe la forma que
    declara P01 (`<tipo>/<slug>` o `<tipo>/<número>-<slug>`), no un texto fijo.
    Sin P01 declarado el documento no cambia."""

    NUMBERED_REGEX = "^(?:(chore|feat|fix|perf|refactor|docs|ci|test)/[0-9]+-[a-z0-9-]+|archive/.+|main)$"
    SLUG_BRANCH_LINE = ("- Nombre de rama: `<tipo>/<slug>`, con el slug del t\u00edtulo de la issue "
                        "en min\u00fasculas y guiones y sin n\u00famero de issue:")
    NUMBERED_BRANCH_LINE = ("- Nombre de rama: `<tipo>/<n\u00famero>-<slug>`, con el n\u00famero de la issue "
                            "y el slug del t\u00edtulo en min\u00fasculas y guiones:")

    def doc_of(self, tmp, p01=None):
        repo = make_docs_repo(tmp)
        if p01 is not None:
            ci_yaml = repo / "ci-pattern.yaml"
            ci_yaml.write_text(
                ci_yaml.read_text(encoding="utf-8")
                + f'P01_branch_name_pattern: "{p01}"\n', encoding="utf-8")
        out = Path(tmp) / "doc.md"
        code, _, err = run_cli(
            ["adoption", "generate-doc", str(repo), "--out", str(out)])
        self.assertEqual(code, 0, err)
        return out.read_text(encoding="utf-8")

    def test_default_p01_keeps_slug_wording(self):
        with tempfile.TemporaryDirectory() as tmp:
            text = self.doc_of(tmp)
        self.assertIn(self.SLUG_BRANCH_LINE, text)
        self.assertIn("  `git switch -c <tipo>/<slug>`", text)
        self.assertIn("  `<rama-nueva>` = `<tipo>/<slug>`, con el `<slug>` del t\u00edtulo de la "
                      "issue en min\u00fasculas y guiones, sin n\u00famero de issue", text)

    def test_numbered_p01_describes_issue_number(self):
        with tempfile.TemporaryDirectory() as tmp:
            text = self.doc_of(tmp, p01=self.NUMBERED_REGEX)
        self.assertIn(self.NUMBERED_BRANCH_LINE, text)
        self.assertIn("  `git switch -c <tipo>/<n\u00famero>-<slug>`", text)
        self.assertIn("  `<rama-nueva>` = `<tipo>/<n\u00famero>-<slug>`, con el n\u00famero de la "
                      "issue", text)
        self.assertIn(self.NUMBERED_REGEX, text)
        self.assertNotIn("sin n\u00famero de issue", text)
        self.assertNotIn("<tipo>/<slug>", text)


class ChecksTableTests(unittest.TestCase):
    """#305: la sección 'required-checks' es una tabla generada 1:1 desde
    los jobs requeridos reales; cada fila con qué verifica, el comando
    local que la reproduce (HR-4) y su receta de recuperación. El gate de
    la fase 5 compara tabla y jobs en AMBOS sentidos."""

    def doc_of(self, tmp, extra_jobs=()):
        repo = make_docs_repo(tmp)
        if extra_jobs:
            policy_path = repo / ".github" / "required-jobs-policy.json"
            policy = json.loads(policy_path.read_text(encoding="utf-8"))
            policy["required_jobs"] += list(extra_jobs)
            policy_path.write_text(json.dumps(policy), encoding="utf-8")
        out = Path(tmp) / "doc.md"
        code, _, err = run_cli(
            ["adoption", "generate-doc", str(repo), "--out", str(out)])
        self.assertEqual(code, 0, err)
        text = out.read_text(encoding="utf-8")
        start = text.index("<!-- ci-pattern-doc:begin required-checks -->")
        end = text.index("<!-- ci-pattern-doc:end required-checks -->")
        return text[start:end]

    def table_rows(self, body):
        rows = {}
        for line in body.splitlines():
            m = re.match(r"^\| `([a-z0-9-]+)` \|", line)
            if m:
                rows[m.group(1)] = line
        return rows

    def test_one_row_per_required_job(self):
        body = self.doc_of(tempfile.mkdtemp())
        rows = self.table_rows(body)
        self.assertEqual(sorted(rows), ["lint", "unit"],
                         "la tabla debe cubrir 1:1 los jobs requeridos")

    def test_every_row_has_command_and_recovery(self):
        body = self.doc_of(tempfile.mkdtemp())
        for name, line in self.table_rows(body).items():
            self.assertIn("bash testing/run-unit-suites.sh", line,
                          f"la fila {name} sin comando local que la reproduzca")
            self.assertIn("→", line,
                          f"la fila {name} sin receta de recuperación")

    def test_recovery_is_check_specific(self):
        body = self.doc_of(tempfile.mkdtemp(), extra_jobs=["pr-size"])
        rows = self.table_rows(body)
        # pr-size: la recuperación manda a reducir el diff o partir en
        # cadena; unit: a leer el primer fallo del preflight.
        self.assertIn("cadena", rows.get("pr-size", ""))
        self.assertIn("fallo", rows.get("unit", ""))

    def test_table_gate_missing_job_is_finding(self):
        # dirección 1: un job requerido sin fila en la tabla del doc (la
        # fila se elimina del doc, no la política: la fase 4 debe seguir
        # en verde para que el hallazgo sea el de la tabla).
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            doc = repo / "docs" / "ci-pattern-flow.md"
            text = doc.read_text(encoding="utf-8")
            line = next(item for item in text.splitlines()
                        if item.startswith("| `unit` |"))
            doc.write_text(text.replace(line + "\n", "", 1), encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("unit", err)
            self.assertIn("no aparece en la tabla", err)

    def test_table_gate_unknown_check_is_finding(self):
        # dirección 2: una fila que documenta un check inexistente.
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            doc = repo / "docs" / "ci-pattern-flow.md"
            text = doc.read_text(encoding="utf-8")
            line = next(item for item in text.splitlines()
                        if item.startswith("| `unit` |"))
            doc.write_text(
                text.replace(line, line + "\n| `ghost-check` | inventado | no | no |", 1),
                encoding="utf-8")
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("ghost-check", err)
            self.assertIn("no existe", err)

class AgentsBlockTests(unittest.TestCase):
    """#305: AGENTS.md del consumer lleva el bloque canónico generado
    (índice de skills por situación + mandato «no inventar, deténgase y
    pregunte»); el gate de la fase 5 verifica el bloque byte a byte."""

    def agents_of(self, tmp):
        repo = make_docs_repo(tmp)
        out = Path(tmp) / "doc.md"
        code, _, err = run_cli(
            ["adoption", "generate-doc", str(repo), "--out", str(out)])
        self.assertEqual(code, 0, err)
        text = out.read_text(encoding="utf-8")
        start = text.index("<!-- ci-pattern-doc:begin agents-block -->")
        end = text.index("<!-- ci-pattern-doc:end agents-block -->")
        return text[start:end]

    def test_doc_emits_canonical_agents_block(self):
        body = self.agents_of(tempfile.mkdtemp())
        # índice por situación, al estilo del modelo de referencia
        for situation in ("Crear una issue", "Crear una rama",
                          "Escribir commits", "Abrir un pull request",
                          "Encadenar PRs", "Consultar checks",
                          "Pedir un merge", "Reportar una fricción"):
            self.assertIn(situation, body, f"falta la situación: {situation}")
        # mandato «no inventar»
        self.assertIn("no lo invente", body)
        self.assertIn("deténgase y pregunte", body)

    def test_gate_requires_block_in_agents_md(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            agents = repo / "AGENTS.md"
            text = agents.read_text(encoding="utf-8")
            begin = "<!-- ci-pattern-agents:begin -->"
            end = "<!-- ci-pattern-agents:end -->"
            agents.write_text(
                text[:text.index(begin)]
                + text[text.index(end) + len(end) + 1:], encoding="utf-8")
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("AGENTS.md", err)
            self.assertIn("no lo invente", err)

    def test_gate_rejects_altered_block(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=5)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            seed_operating_doc(repo)
            agents = repo / "AGENTS.md"
            # una copia alterada del mandato es peor que su ausencia: la IA
            # confiaría en una regla que el patrón no emitió
            agents.write_text(
                agents.read_text(encoding="utf-8")
                .replace("no lo invente", "puede inferirse", 1),
                encoding="utf-8")
            seed_manifest(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("alterado", err)

class RecursionGuardTests(unittest.TestCase):
    """Guarda anti-recursión (#277, incidente 2026-10-05): fase 3 lanzaba
    `ci-pattern verify` como subproceso, verify ejecuta `adoption check` en
    proceso, y el ciclo superó 1.300 procesos python tumbo el VPS por OOM.
    La identidad del manifiesto se comprueba EN PROCESO y la marca
    CI_PATTERN_NESTED hace que cualquier subproceso anidado de verify o
    adoption falle cerrado con exit 2."""

    def test_nested_adoption_is_exit_2(self):
        os.environ["CI_PATTERN_NESTED"] = "1"
        try:
            with tempfile.TemporaryDirectory() as tmp:
                repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
                code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
                self.assertEqual(code, 2)
                self.assertIn("recursión de ci-pattern detectada", err)
        finally:
            os.environ.pop("CI_PATTERN_NESTED", None)

    def test_nested_verify_is_exit_2(self):
        os.environ["CI_PATTERN_NESTED"] = "1"
        try:
            with tempfile.TemporaryDirectory() as tmp:
                repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
                code, _, err = run_cli(["verify", str(repo)])
                self.assertEqual(code, 2)
                self.assertIn("recursión de ci-pattern detectada", err)
        finally:
            os.environ.pop("CI_PATTERN_NESTED", None)

    def test_without_mark_there_is_no_exit_2(self):
        # Sin la marca el veredicto es el de siempre (no la guarda): la
        # guarda solo dispara sobre subprocesos anidados reales.
        os.environ.pop("CI_PATTERN_NESTED", None)
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=0))
            code, _, err = run_cli(["adoption", "check", "--all", str(repo)])
            self.assertNotEqual(code, 2)
            self.assertNotIn("recursión de ci-pattern detectada", err)

    def test_child_env_marks_nesting(self):
        env = CLI._child_env()
        self.assertEqual(env.get("CI_PATTERN_NESTED"), "1")

    def test_every_subprocess_run_has_timeout(self):
        # Regla permanente: todo subprocess.run en la CLI lleva timeout=.
        # Meta-test estático: si alguien añade un subprocess.run sin
        # timeout, esta prueba falla antes de que un gate pueda colgar.
        source = CLI_PATH.read_text(encoding="utf-8")
        for match in re.finditer(r"subprocess\.run\(", source):
            depth, i = 1, match.end()
            while depth and i < len(source):
                if source[i] == "(":
                    depth += 1
                elif source[i] == ")":
                    depth -= 1
                i += 1
            call = source[match.start():i]
            self.assertIn("timeout=", call,
                          f"subprocess.run sin timeout=: {call[:120]!r}")

    def test_phase_three_identity_is_in_process(self):
        # La fase 3 NO lanza verify ni adoption como subproceso: el código
        # del gate no contiene el patrón de subproceso del propio binario.
        source = CLI_PATH.read_text(encoding="utf-8")
        phase_three = source.split("def _check_phase_three", 1)[1].split(
            "\ndef ", 1)[0]
        for forbidden in ("ci-pattern", "verify", "adoption"):
            self.assertNotIn(
                f'"{forbidden}', phase_three,
                "la fase 3 no debe invocar subcomandos ci-pattern como subproceso")
        self.assertIn("_manifest_identity_findings", phase_three)

# #513 / #540: PhaseFourSmokeBoundedTests y PHASE_FOUR_BAIT_TEST se retiraron
# con el meta-gate y la matriz HR→gate. La guarda de profundidad que
# verificaban (#342) ya no es necesaria: sin meta-gate, ningún gate del
# patrón ejecuta el test declarado por una regla HR, así que la cadena
# CLI → meta-gate → CLI nunca se dispara. Si volviera a haber un meta-gate,
# reintroducir también este test con la misma sonda.

PLAN_403_SNAPSHOT = {
    "message": "Upgrade to GitHub Pro or make this repository public to enable this feature.",
    "documentation_url": "https://docs.github.com/rest", "status": "403"}
RUNNER_LABEL = "cadete"


class RunnerEnforcedAdoptionTests(unittest.TestCase):
    """#303: el camino runner de la adopción (fases 3 y 4) sobre un host sin
    protección de rama: compensatorios ausentes (exit 1 nombrando lo que
    falta) y presentes (exit 0); runner online según actions/runners."""

    def private_contract(self, upto=4, with_binding_gate=True):
        contract = complete_contract(upto=upto)
        if upto >= 3:
            contract["phases"]["3"]["host_snapshots"]["actions-runners"] = \
                "docs/adoption/actions-runners.json"
        if upto >= 4 and with_binding_gate:
            contract["phases"]["4"] = full_phase4_gates(extra=[{
                "gate": "runner-binding",
                "args": ["--contract", "@.github/host-contract.json",
                         "--workflows", "@.github/workflows"]}])
        return contract

    def private_host_contract(self):
        return {
            "contract_version": 1,
            "labels": [{"name": "status:approved", "class": "host-enforced"},
                       {"name": "chain:partial", "class": "host-enforced"}],
            "required_checks": [{"name": "branch-name", "class": "runner-enforced"}],
            "merge_methods": {"allow_merge_commit": {"declared": True, "class": "host-enforced"}},
            "protection": {"enforce_admins": {"declared": False, "class": "runner-enforced"}},
            "rulesets": [],
            "host_capabilities": {"branch_protection": "unavailable",
                                  "rulesets": "unavailable", "required_checks": "unavailable"},
            "runner": {"label": RUNNER_LABEL},
            "protected_branches": ["main"],
        }

    def seed_private_host(self, repo, online=True):
        (repo / ".github" / "host-contract.json").write_text(
            json.dumps(self.private_host_contract()), encoding="utf-8")
        # el contrato del host es artefacto sellado: se re-sella tras el cambio
        seed_manifest(repo, list(PATTERN_MANIFEST_FILES))
        touch(repo, "docs/adoption/branch-protection.json", json.dumps(PLAN_403_SNAPSHOT))
        touch(repo, "docs/adoption/rulesets.json", json.dumps(PLAN_403_SNAPSHOT))
        touch(repo, "docs/adoption/actions-runners.json", json.dumps({
            "total_count": 1,
            "runners": [{"name": "cadete-oracle-arm64",
                         "status": "online" if online else "offline",
                         "labels": [{"name": "self-hosted"}, {"name": RUNNER_LABEL}]}]}))
        touch(repo, "api/repo.json", json.dumps({"allow_merge_commit": True}))
        touch(repo, "api/labels.json", json.dumps(
            [{"name": "status:approved"}, {"name": "chain:partial"}]))
        touch(repo, "api/branch-protection.json", json.dumps(PLAN_403_SNAPSHOT))
        touch(repo, "api/rulesets.json", json.dumps(PLAN_403_SNAPSHOT))

    def wire_workflows(self, repo):
        touch(repo, ".github/workflows/guard.yml",
              "name: guard\non: push\njobs:\n  guard:\n"
              f"    runs-on: [self-hosted, {RUNNER_LABEL}]\n    steps:\n"
              "      - run: python3 x/check_force_push_guard.py\n"
              "      - run: python3 x/check_push_compliance.py\n")
        touch(repo, ".github/workflows/merge.yml",
              "name: merge\non: workflow_dispatch\njobs:\n  merge:\n"
              f"    runs-on: [self-hosted, {RUNNER_LABEL}]\n    steps:\n"
              "      - run: python3 x/check_governed_merge.py\n")
        # los workflows nuevos son gobierno del consumer: inventariados (fase 0)
        # y el contrato re-sellado tras el cambio.
        contract = json.loads(
            (repo / ".github" / "ci-pattern-adoption.json").read_text(encoding="utf-8"))
        inventory = contract["phases"]["0"]["inventory"]
        for rel in (".github/workflows/guard.yml", ".github/workflows/merge.yml"):
            if not any(isinstance(e, dict) and e.get("artifact") == rel for e in inventory):
                inventory.append({"artifact": rel, "disposition": "adopted", "target": rel})
        write_contract(repo, contract)

    def test_phase_three_requires_actions_runners_snapshot_on_runner_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = complete_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("actions-runners", err)

    def test_phase_three_runner_path_with_plan403_is_green(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.private_contract(upto=3)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo)
            code, out, err = run_cli(["adoption", "check", "--phase", "3", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_phase_4_runner_path_without_binding_gate_is_a_violation(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.private_contract(with_binding_gate=False)
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("runner-binding", err)

    def test_phase_4_unwired_workflows_are_a_violation(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.private_contract()
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("runner-binding", err)
            self.assertIn("UNBOUND", err)

    def test_phase_4_offline_runner_is_a_violation(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.private_contract()
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo, online=False)
            self.wire_workflows(repo)
            code, _, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("online", err)

    def test_phase_4_runner_path_fully_green_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = self.private_contract()
            repo, _ = make_contract_repo(tmp, contract=contract)
            seed_evidence(repo, 3)
            self.seed_private_host(repo)
            self.wire_workflows(repo)
            code, out, err = run_cli(["adoption", "check", "--phase", "4", str(repo)])
            self.assertEqual(code, 0, out + err)


class AdoptionContractSchemaTests(unittest.TestCase):
    """#369 + #606: el esquema publicado y la CLI del mismo asset admiten
    las mismas claves de fase (claves 0..ADOPTION_PHASES-1) y los mismos
    nombres de gate en la fase 4. La CLI no lee el esquema, así que sin
    estos tests la deriva entre los dos pasa inadvertida y un consumer
    (p. ej. CADETE, #606) queda bloqueado con `gate desconocido en el
    contrato` aunque su contrato sea semánticamente válido."""

    SCHEMA_PATH = SKILL_ROOT / "assets" / "adoption" / "adoption-contract.schema.json"

    def phase_pattern(self):
        schema = json.loads(self.SCHEMA_PATH.read_text(encoding="utf-8"))
        return schema["properties"]["phases"]["propertyNames"]["pattern"]

    def phase4_comment(self):
        schema = json.loads(self.SCHEMA_PATH.read_text(encoding="utf-8"))
        return schema["$defs"]["phase"]["$comment"]

    def phase4_gates(self):
        """Extrae del `$comment` de `phase` el conjunto de gates que el
        esquema nombra para la fase 4: la lista explícita dentro de la
        cláusula «cubriendo EXACTAMENTE el conjunto conocido (...)» y la
        condicional «+ X cuando el host declara camino runner». El
        anclaje en la sub-frase única de la fase 4 es deliberado: la
        fase 3 también usa «+ X cuando el host declara camino runner»
        (con `actions-runners`), y un parseo ingenuo se la llevaría
        puesta. Si la sub-frase cambia, este test rompe y obliga a
        actualizar el parseo junto con la documentación — es la
        invariante que protege contra la regresión #606."""
        comment = self.phase4_comment()
        anchor = ("fase 4 exige gates cubriendo EXACTAMENTE el conjunto "
                  "conocido (")
        start = comment.find(anchor)
        if start < 0:
            self.fail(
                "$comment no contiene la cláusula 'fase 4 ... conjunto "
                "conocido (...)' que el regresor analiza; actualice el "
                "anchor de `phase4_gates` junto con el texto del "
                "$comment")
        start += len(anchor)
        end = comment.find(")", start)
        if end < 0:
            self.fail("$comment: la cláusula de fase 4 no cierra paréntesis")
        body = comment[start:end]
        chunks = body.split(";", 1)
        base = chunks[0].strip()
        base_gates = [g.strip() for g in base.split("/") if g.strip()]
        conditional = []
        if len(chunks) > 1:
            m = re.search(
                r"\+\s*([\w-]+)\s+cuando el host declara camino runner",
                chunks[1])
            if m:
                conditional = [m.group(1)]
        return set(base_gates) | set(conditional)

    def test_schema_phase_keys_match_the_cli(self):
        pattern = self.phase_pattern()
        for key in (str(n) for n in range(CLI.ADOPTION_PHASES)):
            self.assertTrue(
                re.fullmatch(pattern, key),
                f"el esquema no admite la clave de fase '{key}' que la CLI sí admite")
        self.assertIsNone(
            re.fullmatch(pattern, str(CLI.ADOPTION_PHASES)),
            "el esquema admite una clave de fase que la CLI no admite")

    def test_schema_phase4_gate_set_matches_the_cli(self):
        """#606 gemelo de #369 para los nombres de gate del `$comment`:
        el conjunto que el esquema enumera para la fase 4 debe coincidir
        exactamente con `ADOPTION_GATES ∪ RUNNER_BINDING_GATES` del
        binario `assets/bin/ci-pattern`. Si el esquema nombra un gate
        que la CLI no conoce, el contrato del consumer lo acepta y la
        CLI lo rechaza con `gate desconocido en el contrato` (#606,
        CADETE). Si la CLI añade un gate y el `$comment` no lo nombra,
        el contrato del consumer aceptará cualquier coincidencia parcial
        y la auditoría de cobertura será ciega al hueco."""
        mentioned = self.phase4_gates()
        known = set(CLI.ADOPTION_GATES) | set(CLI.RUNNER_BINDING_GATES)
        self.assertTrue(
            mentioned.issubset(known),
            f"el $comment nombra gates que la CLI no conoce: "
            f"{sorted(mentioned - known)}; "
            f"registro vivo de la CLI: {sorted(known)}")
        self.assertTrue(
            known.issubset(mentioned),
            f"el $comment omite gates que la CLI sí conoce: "
            f"{sorted(known - mentioned)}; "
            f"mencionados en el $comment: {sorted(mentioned)}")

    def test_schema_phase4_comment_does_not_name_retired_hr_matrix(self):
        """#606 defensa específica contra la regresión original: el
        `$comment` no nombra `hr-matrix` como gate del conjunto
        conocido. La retirada del meta-gate por #513/#540 lo deja fuera
        del registro vivo; reintroducirlo requiere reintroducir la regla
        y su `CI-enforcement`, no basta con añadir el nombre al
        `$comment`."""
        self.assertNotIn(
            "hr-matrix", self.phase4_comment(),
            "el $comment cita 'hr-matrix' como gate de la fase 4, pero "
            "la CLI lo retiró por #513/#540; reintroducir requiere "
            "reintroducir la regla completa, no basta con citar el "
            "nombre")


class PhaseZeroWorkflowCommentsTests(unittest.TestCase):
    """#607: la fase 0 no cuenta como gobierno las rutas que solo viven en
    un comentario YAML de un workflow. Solo cuentan las referenciadas por
    claves ejecutables (`run:`, `uses:`, `with:`, `paths:`). El bug actual
    exige inventario para cualquier path que case con las extensiones del
    catálogo, sin distinguir entre una referencia «ejecutable» y una
    mención en prosa dentro de un comentario `# ...`."""

    def _build_repo(self, *, workflow_body, paths_with_files):
        """Construye un repo mínimo con un workflow y los archivos que el
        bug detecta por referencia. Devuelve la ruta del repo.

        - `workflow_body`: contenido del workflow en ``.github/workflows/ci.yml``.
        - `paths_with_files`: lista de rutas relativas (str) que el fixture
          debe crear vacías en el árbol; el bug las exige cuando aparecen
          en el texto del workflow, exista el archivo o no, pero el filtro
          de existencia de ``_scan_governed_paths`` exige que el archivo
          exista para que la falsa exigencia llegue al gate."""
        d = tempfile.mkdtemp(prefix="ci-pattern-607-")
        # mktemp no borra el dir al salir: la regla del operador
        # («no borres temporales») deja el árbol bajo /tmp/ci-pattern-607-*;
        # el sistema limpia /tmp por su cuenta.
        repo = Path(d) / "repo"
        (repo / ".github" / "workflows").mkdir(parents=True)
        (repo / ".github" / "workflows" / "ci.yml").write_text(
            workflow_body, encoding="utf-8")
        (repo / ".github" / "host-labels-snapshot.json").write_text(
            "[]", encoding="utf-8")
        (repo / ".github" / "ci-pattern-adoption.json").write_text(json.dumps({
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": ".github/workflows/ci.yml",
                         "disposition": "adopted",
                         "target": ".github/workflows/ci.yml"},
                    ],
                    "host_labels": ".github/host-labels-snapshot.json",
                    "removal_confirmed": True,
                }
            },
        }, indent=2, ensure_ascii=False), encoding="utf-8")
        for rel in paths_with_files:
            target = repo / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            if not target.exists():
                target.write_text("", encoding="utf-8")
        return repo

    def test_path_in_yaml_comment_is_not_a_governed_artifact(self):
        """#607 S1: una ruta que solo aparece en un comentario YAML del
        workflow NO exige entrada de inventario en la fase 0. Antes del
        fix, la fase 0 marca la ruta como «artefacto de gobierno sin
        entrada de inventario» y sale con ``exit=1``; después del fix,
        la fase 0 no ve la ruta y sale con ``exit=0``."""
        repo = self._build_repo(
            workflow_body=(
                "name: ci\n"
                "on: [push]\n"
                "jobs:\n"
                "  test:\n"
                "    runs-on: ubuntu-latest\n"
                "    steps:\n"
                # La ruta está SOLO en este comentario: el run: de abajo
                # no la invoca. El bug actual exige inventario igualmente.
                "      # this references app/migrations/env.py but it is just a comment\n"
                "      # and tests/lanzadera/migrations/test_seed_migrations.py too\n"
                "      - run: echo hello\n"
            ),
            paths_with_files=[
                "app/migrations/env.py",
                "tests/lanzadera/migrations/test_seed_migrations.py",
            ],
        )
        code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
        joined = out + err
        # La fase 0 no debe exigir las rutas en comentario.
        self.assertNotIn(
            "app/migrations/env.py", joined,
            f"la fase 0 exige app/migrations/env.py aunque solo aparece en "
            f"un comentario YAML; salida del CLI:\n{joined}")
        self.assertNotIn(
            "tests/lanzadera/migrations/test_seed_migrations.py", joined,
            f"la fase 0 exige tests/lanzadera/migrations/test_seed_migrations.py "
            f"aunque solo aparece en un comentario YAML; salida del CLI:\n{joined}")
        # Y la fase 0 debe pasar limpia (rc=0): no debe fallar por otra razón
        # (cambio de contrato, cobertura de etiquetas, etc.).
        self.assertEqual(
            code, 0,
            f"fase 0 en verde: ningún hallazgo para paths en comentarios YAML; "
            f"salida del CLI:\n{joined}")

    def test_path_in_run_is_a_governed_artifact(self):
        """#607 S2 (regresión negativa): una ruta invocada en ``run:`` SÍ
        exige entrada de inventario. El fix no debe debilitar este caso
        ni eliminar el hallazgo que ya era correcto."""
        repo = self._build_repo(
            workflow_body=(
                "name: ci\n"
                "on: [push]\n"
                "jobs:\n"
                "  test:\n"
                "    runs-on: ubuntu-latest\n"
                "    steps:\n"
                "      - run: bash scripts/coverage-ratchet.sh\n"
            ),
            paths_with_files=["scripts/coverage-ratchet.sh"],
        )
        code, out, err = run_cli(["adoption", "check", "--phase", "0", str(repo)])
        joined = out + err
        # El path en run: SÍ debe seguir exigiéndose.
        self.assertIn(
            "scripts/coverage-ratchet.sh", joined,
            f"la fase 0 debe exigir scripts/coverage-ratchet.sh cuando "
            f"aparece en run:; salida del CLI:\n{joined}")
        self.assertEqual(
            code, 1,
            f"fase 0 sigue exigiendo paths referenciados en run: "
            f"(regresión negativa del fix #607); salida del CLI:\n{joined}")


HR44_LINE = "Los fragmentos de `changelog.d/` se consolidan en el PR de release"
GATE_WORKFLOW = (
    "on:\n  pull_request:\njobs:\n  changelog:\n    steps:\n"
    "      - run: python3 personal/ardelperal/ci-pattern/assets/changelog-fragment/"
    "check_changelog_fragment.py --base origin/main --head HEAD\n")


class ChangelogFragmentLineTests(unittest.TestCase):
    """HR-44 condicional (#875, oleada 2 de #873): la línea del doc operativo
    sólo aparece si algún .github/workflows/*.yml cablea el gate
    check_changelog_fragment.py. Sin campo nuevo en el host-contract."""

    def generate(self, tmp, wire_gate):
        repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=5))
        seed_evidence(repo, 3)
        if wire_gate:
            (repo / ".github" / "workflows" / "tests.yml").write_text(
                GATE_WORKFLOW, encoding="utf-8")
        seed_operating_doc(repo)
        return repo, (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")

    def test_hr44_line_absent_without_gate_wiring(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, doc = self.generate(tmp, wire_gate=False)
            self.assertNotIn(HR44_LINE, doc)

    def test_hr44_line_present_with_gate_wiring(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, doc = self.generate(tmp, wire_gate=True)
            self.assertEqual(doc.count(HR44_LINE), 1)

    def test_removing_gate_wiring_is_drift_in_pull_request_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = self.generate(tmp, wire_gate=True)
            (repo / ".github" / "workflows" / "tests.yml").unlink()
            code, _, err = run_cli(["adoption", "check", "--phase", "5", str(repo)])
            self.assertEqual(code, 1, err)
            self.assertIn("deriva", err)
            self.assertIn("pull-request", err)

    def test_hr44_line_present_with_yaml_extension_wiring(self):
        # #883: GitHub acepta .github/workflows/*.yaml además de *.yml;
        # el helper debe reconocer ambas extensiones.
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_contract_repo(tmp, contract=complete_contract(upto=5))
            seed_evidence(repo, 3)
            (repo / ".github" / "workflows" / "tests.yaml").write_text(
                GATE_WORKFLOW, encoding="utf-8")
            seed_operating_doc(repo)
            doc = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            self.assertEqual(doc.count(HR44_LINE), 1)


if __name__ == "__main__":
    unittest.main()
