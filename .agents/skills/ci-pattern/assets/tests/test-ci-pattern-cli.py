#!/usr/bin/env python3
"""Suite de tests de la CLI `ci-pattern` (assets/tests de la skill).

Ejecución: `python3 test-ci-pattern-cli.py` — stdlib exclusivamente, sin
instalación, apta para CI. El fichero de fixtures se construye en un directorio
temporal por test (repo destino falso); ningún test escribe fuera de `tmp`.

Cubre: validación de parámetros (válida + 4 formas inválidas), round-trip del
manifiesto, verify limpio, verify con drift (fichero modificado, gate sin
cablear, bloque de slice desfasado), status, stubs adopt/update.
"""

import contextlib
import hashlib
import importlib.util
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[2]
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"
SKILL_DIR = SKILL_ROOT


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli", importlib.machinery.SourceFileLoader("ci_pattern_cli", str(CLI_PATH))
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()

VALID_PARAMS = """\
# parámetros de ejemplo
P06_review_budget_lines: 400
P10_label_approval: status:approved
P18_health_url_var: MY_DEPLOY_HEALTH_URL
P21_smoke_user_agent: my-smoke/1
P45_deploy_image: ghcr.io/example/app
"""

INVALID = {
    "unknown_key": "P99_nonexistent: 1\n" + VALID_PARAMS,
    "missing_required": "P06_review_budget_lines: 400\nP10_label_approval: status:approved\n",
    "wrong_type": VALID_PARAMS.replace("P06_review_budget_lines: 400", 'P06_review_budget_lines: "cuatrocientos"'),
    "bad_regex": VALID_PARAMS + 'P01_branch_name_pattern: "([unclosed"\n',
    "bad_label": VALID_PARAMS.replace("status:approved", "approved"),
    "bad_path": VALID_PARAMS + "P41_gate_policy_path: /etc/passwd\n",
}


def sha(path):
    data = Path(path).read_bytes()
    return "sha256:" + hashlib.sha256(data).hexdigest()


def write_slice_block(repo, name, content, version="v1"):
    block = (
        f"<!-- personal-skills:slice:{name} @ {version} -->\n"
        f"{content}\n"
        f"<!-- /personal-skills:slice:{name} -->\n"
    )
    agents = Path(repo) / "AGENTS.md"
    agents.write_text(("previo\n\n" if agents.exists() else "") + block, encoding="utf-8")
    return "sha256:" + hashlib.sha256((content + "\n").encode("utf-8")).hexdigest()


def make_repo(tmp, params=None, jobs=("lint",), extra_jobs=(), slice_name="APAP_WEB", files=("scripts/check_pr_size.py",)):
    repo = Path(tmp) / "repo"
    wf = repo / ".github" / "workflows"
    wf.mkdir(parents=True)
    (wf / "ci.yml").write_text(
        "jobs:\n"
        + "".join(f"  {j}:\n    runs-on: ubuntu-24.04\n" for j in jobs),
        encoding="utf-8",
    )
    hashes = {}
    for rel in files:
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(f"# {rel}\n", encoding="utf-8")
        hashes[rel] = sha(p)
    slice_hash = write_slice_block(repo, slice_name, "contenido del slice\nsegunda línea\n")
    params_path = repo / "ci-pattern.yaml"
    params_path.write_text(params if params is not None else VALID_PARAMS, encoding="utf-8")
    manifest = {
        "schema_version": 1,
        "adopted_at": "2026-10-02T12:00:00Z",
        "canonical_version": "0.4",
        "canonical_source": {"repo": "DysTelefonica/team-skills", "path": "personal/ardelperal/ci-pattern"},
        "parameters": {"P23_required_jobs": list(jobs) + list(extra_jobs)},
        "files": hashes,
        "slice_blocks": {slice_name: slice_hash},
    }
    (repo / ".governance-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return repo, params_path


def run_cli(argv, env_extra=None):
    out, err = io.StringIO(), io.StringIO()
    env = dict(os.environ)
    env["CI_PATTERN_SKILL_DIR"] = str(SKILL_DIR)
    if env_extra:
        env.update(env_extra)
    old = os.environ.copy()
    os.environ.clear()
    os.environ.update(env)
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = CLI.main(argv)
    finally:
        os.environ.clear()
        os.environ.update(old)
    return code, out.getvalue(), err.getvalue()


class ParamsValidateTests(unittest.TestCase):
    def test_valid_file_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp)
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 0, out)

    def test_unknown_key_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["unknown_key"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P99_nonexistent", out)

    def test_missing_required_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["missing_required"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P18_health_url_var", out)

    def test_wrong_type_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["wrong_type"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P06_review_budget_lines", out)
            self.assertIn("integer", out)

    def test_bad_regex_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["bad_regex"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P01_branch_name_pattern", out)

    def test_bad_label_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["bad_label"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P10_label_approval", out)
            self.assertIn("prefijo:valor", out)

    def test_absolute_path_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=INVALID["bad_path"])
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P41_gate_policy_path", out)

    def test_missing_file_is_exit_3(self):
        code, _, _ = run_cli(["params", "validate", "/nonexistent/ci-pattern.yaml"])
        self.assertEqual(code, 3)

    def test_empty_map_is_accepted(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=VALID_PARAMS + "P09_labels_type: {}\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 0, out)

    def test_inline_map_with_entries_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = make_repo(tmp, params=VALID_PARAMS + "P50_secrets_manager: {a: b}\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1, out)
            self.assertIn("P50_secrets_manager", out)
            self.assertIn("mapa inline", out)


class ManifestRoundTripTests(unittest.TestCase):
    def test_roundtrip_preserves_fields(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            loaded = CLI.load_manifest(repo)
            raw = json.loads((repo / ".governance-manifest.json").read_text())
            self.assertEqual(loaded, raw)
            self.assertTrue(all(v.startswith("sha256:") for v in loaded["files"].values()))
            self.assertTrue(all(v.startswith("sha256:") for v in loaded["slice_blocks"].values()))

    def test_write_manifest_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo2"
            repo.mkdir()
            manifest = {
                "schema_version": 1,
                "adopted_at": "2026-10-02T00:00:00Z",
                "canonical_version": "0.4",
                "canonical_source": {"repo": "r", "path": "p"},
                "parameters": {},
                "files": {"a.txt": "sha256:" + "0" * 64},
                "slice_blocks": {},
            }
            CLI.write_manifest(repo, manifest)
            self.assertEqual(CLI.load_manifest(repo), manifest)


class VerifyTests(unittest.TestCase):
    def test_clean_repo_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out)
            self.assertIn("limpio", out)

    def test_modified_file_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            (repo / "scripts" / "check_pr_size.py").write_text("# editado localmente\n")
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("locally-modified", out)
            self.assertIn("check_pr_size.py", out)

    def test_missing_gate_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp, extra_jobs=["pr-size"])  # workflows solo tienen lint
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("gate-not-wired", out)
            self.assertIn("pr-size", out)

    def test_stale_slice_block_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            write_slice_block(repo, "APAP_WEB", "contenido EDITADO localmente\n")
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("locally-modified", out)
            self.assertIn("APAP_WEB", out)

    def test_stale_vs_canonical_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            partials = canon / "partials"
            partials.mkdir(parents=True)
            (partials / "web.md").write_text("# Web\ncontenido partial web\n", encoding="utf-8")
            frag_dir = canon / "APAP_WEB"
            frag_dir.mkdir()
            (frag_dir / "AGENTS.fragment.md").write_text("fragment APAP_WEB\n", encoding="utf-8")
            fleet = Path(tmp) / "fleet"
            fleet.mkdir()
            (fleet / "registry.json").write_text(
                json.dumps({"schema_version": 1, "repositories": [
                    {"name": "APAP_WEB", "canonical_name": "APAP_WEB", "primary_type": "web"}
                ]}),
                encoding="utf-8",
            )
            repo, _ = make_repo(tmp)
            # el consumer edita el cuerpo tras propagar y ajusta el manifiesto
            # a su edición: sin locally-modified, pero el bloque difiere de la
            # composición canónica y su marcador ya no es la versión vigente.
            edited = "# Web\ncontenido partial web\n\n\nfragment APAP_WEB editado a mano\n"
            sha7 = hashlib.sha256(edited.rstrip().encode("utf-8")).hexdigest()[:7]
            agents = repo / "AGENTS.md"
            agents.write_text(
                f"<!-- personal-skills:slice:APAP_WEB @ v{sha7} -->\n{edited.rstrip()}\n<!-- /personal-skills:slice:APAP_WEB -->\n",
                encoding="utf-8",
            )
            manifest = json.loads((repo / ".governance-manifest.json").read_text(encoding="utf-8"))
            manifest["slice_blocks"]["APAP_WEB"] = "sha256:" + hashlib.sha256((edited.rstrip() + "\n").encode("utf-8")).hexdigest()
            (repo / ".governance-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", canon])
            self.assertEqual(code, 1)
            self.assertIn("stale-vs-canonical", out)

    def test_json_output_shape(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, _ = run_cli(["verify", str(repo), "--json"])
            payload = json.loads(out)
            self.assertEqual(code, 0)
            self.assertTrue(payload["clean"])
            self.assertEqual(payload["findings"], [])
            self.assertEqual(payload["canonical_comparison"], "skipped")

    def test_missing_manifest_exit_3(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "vacio"
            repo.mkdir()
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 3)
            self.assertIn("adopt", out)


class StatusAndStubsTests(unittest.TestCase):
    def test_status_with_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, _ = run_cli(["status", str(repo)])
            self.assertEqual(code, 0, out)
            self.assertIn("files", out)

    def test_status_without_manifest_exit_3(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "vacio"
            repo.mkdir()
            code, _, _ = run_cli(["status", str(repo)])
            self.assertEqual(code, 3)

    def test_adopt_redirects_to_adoption_check(self):
        # #276: sin exit 4 colgante; la adopción la gobierna `adoption check`.
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, err = run_cli(["adopt", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("adoption check", out + err)

    def test_update_redirects_to_adoption_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, err = run_cli(["update", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("adoption check", out + err)

    def test_usage_error_exit_2(self):
        code, _, _ = run_cli(["comando-inexistente"])
        self.assertEqual(code, 2)


class VerifyAdoptionTests(unittest.TestCase):
    """#276: `verify` exige `adoption check --all` en verde cuando el
    consumer declara contrato de adopción; sin contrato, verify no cambia."""

    def make_verified_repo(self, tmp):
        repo = Path(tmp) / "consumer"
        (repo / ".github").mkdir(parents=True)
        (repo / "AGENTS.md").write_text("# consumer\n", encoding="utf-8")
        digest = hashlib.sha256(b"# consumer\n").hexdigest()
        (repo / ".governance-manifest.json").write_text(
            json.dumps({"schema_version": 1, "parameters": {"P23_required_jobs": []}, "files": {"AGENTS.md": "sha256:" + digest}}),
            encoding="utf-8",
        )
        return repo

    def green_contract(self):
        return {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": "legacy", "disposition": "retired"},
                        {"artifact": ".github/workflows/ci.yml", "disposition": "adopted", "target": ".github/workflows/ci.yml"},
                        {"artifact": "AGENTS.md", "disposition": "adopted", "target": "AGENTS.md"},
                        {"artifact": ".github/host-contract.json", "disposition": "adopted", "target": ".github/host-contract.json"},
                        {"artifact": ".github/required-jobs-policy.json", "disposition": "adopted", "target": ".github/required-jobs-policy.json"},
                        {"artifact": ".github/env-isolation-allowlist.json", "disposition": "adopted", "target": ".github/env-isolation-allowlist.json"},
                        {"artifact": ".github/local-ci-parity-policy.json", "disposition": "adopted", "target": ".github/local-ci-parity-policy.json"},
                        {"artifact": "testing/run-unit-suites.sh", "disposition": "adopted", "target": "testing/run-unit-suites.sh"},
                    ],
                    "removal_confirmed": True,
                    "host_labels": ".github/host-labels-snapshot.json",
                },
                "1": {"evidence": {"path": "docs/adoption/propagation-evidence.md"}},
                "2": {
                    "env_isolation": {
                        "root": "tests",
                        "allowlist": ".github/env-isolation-allowlist.json",
                    }
                },
                "3": {
                    "evidence": {"path": "docs/adoption/governance-replacement-pr.md"},
                    "pattern_paths": [".github/workflows/ci.yml", ".github/host-contract.json"],
                    "host_snapshots": {
                        "branch-protection": "docs/adoption/branch-protection.json",
                        "rulesets": "docs/adoption/rulesets.json",
                    },
                },
                "4": {
                    "gates": [
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
                        {"gate": "hr-matrix", "args": ["@skill"]},
                    ]
                },
                "5": {"operating_doc": {"doc": "docs/ci-pattern-flow.md",
                                         "linked_from": "AGENTS.md"}},
                "6": {"prs": [{"number": 101, "sha": "b" * 40, "conclusion": "success"}]},
            },
        }

    def seed_green_tree(self, repo):
        (repo / "AGENTS.md").write_text("# consumer\n", encoding="utf-8")
        (repo / ".github" / "host-labels-snapshot.json").write_text(
            "[]", encoding="utf-8")
        (repo / "ci-pattern.yaml").write_text(
            "P09_labels_type:\n"
            "  bug_report.yml: type:bug\n"
            "  feature_request.yml: type:feature\n", encoding="utf-8")
        (repo / ".github" / "workflows").mkdir(parents=True, exist_ok=True)
        (repo / ".github" / "workflows" / "ci.yml").write_text(
            "on:\n  pull_request:\njobs:\n  unit:\n    steps:\n"
            "      - run: bash testing/run-unit-suites.sh\n      - uses: ./local@8f4b7f8d\n",
            encoding="utf-8")
        (repo / ".github" / "required-jobs-policy.json").write_text(json.dumps({
            "required_jobs": ["unit"],
            "events": {e: {"accepted_skips": []} for e in
                       ("pull_request", "push", "schedule", "workflow_dispatch")},
        }), encoding="utf-8")
        (repo / ".github" / "env-isolation-allowlist.json").write_text(
            json.dumps({"env_vars": []}), encoding="utf-8")
        (repo / "tests").mkdir(exist_ok=True)
        (repo / "tests" / "test_smoke.sh").write_text("#!/bin/sh\necho ok\n", encoding="utf-8")
        for rel in (
            "docs/adoption/propagation-evidence.md",
            "docs/adoption/governance-replacement-pr.md",
            "docs/adoption/env-isolation-green.md",
            "docs/adoption/governance-replacement-pr.md",
        ):
            path = repo / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("evidencia\n", encoding="utf-8")
        (repo / ".github" / "needs").mkdir(parents=True, exist_ok=True)
        (repo / ".github" / "needs" / "needs.json").write_text(
            json.dumps({"unit": {"result": "success"}}), encoding="utf-8")
        (repo / ".github" / "local-ci-parity-policy.json").write_text(json.dumps({
            "ci": {"workflow": ".github/workflows/ci.yml", "job": "unit"},
            "local": {"entrypoint": "testing/run-unit-suites.sh"}, "exclusions": [],
        }), encoding="utf-8")
        (repo / "testing").mkdir(exist_ok=True)
        (repo / "testing" / "run-unit-suites.sh").write_text(
            "#!/bin/sh\nbash testing/run-unit-suites.sh\n", encoding="utf-8")
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
        (repo / "api").mkdir(exist_ok=True)
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
        (repo / "skill" / "assets" / "gate" / "tests").mkdir(parents=True, exist_ok=True)
        (repo / "skill" / "references").mkdir(exist_ok=True)
        (repo / "skill" / "SKILL.md").write_text(
            "---\nname: mini\n---\n\n- **HR-1 — regla**\n", encoding="utf-8")
        (repo / "skill" / "references" / "hr-gate-matrix.json").write_text(json.dumps(
            {"rules": [{"id": "HR-1", "title": "regla", "enforcement": "gate",
                        "asset": "assets/gate/gate.py",
                        "test": "assets/gate/tests/test_gate.py"}]}), encoding="utf-8")
        (repo / "skill" / "assets" / "gate" / "gate.py").write_text("# gate\n", encoding="utf-8")
        (repo / "skill" / "assets" / "gate" / "tests" / "test_gate.py").write_text(
            "print('ok')\n", encoding="utf-8")
        (repo / "docs/adoption/branch-protection.json").write_text(
            json.dumps({"required_status_checks": {"enabled": True}}), encoding="utf-8"
        )
        (repo / "docs/adoption/rulesets.json").write_text(
            json.dumps([{"name": "branch-protection", "enforcement": "active", "class": "branch"}]),
            encoding="utf-8",
        )


    def seal_manifest(self, repo):
        """Re-sella el manifiesto de gobierno sobre los artefactos que el
        patrón instala (identidad verificada por sha256 en la fase 3)."""
        files = {}
        for rel in (
            "AGENTS.md",
            ".github/workflows/ci.yml",
            ".github/host-contract.json",
            ".github/ci-pattern-adoption.json",
            ".github/required-jobs-policy.json",
            ".github/env-isolation-allowlist.json",
            ".github/local-ci-parity-policy.json",
        ):
            data = (repo / rel).read_bytes().replace(b"\r\n", b"\n")
            files[rel] = "sha256:" + hashlib.sha256(data).hexdigest()
        (repo / ".governance-manifest.json").write_text(
            json.dumps({"schema_version": 1,
                        "parameters": {"P23_required_jobs": []},
                        "files": files}),
            encoding="utf-8",
        )

    def init_git(self, repo):
        subprocess.run(["git", "-C", str(repo), "init", "-q", "-b", "main"], check=True)
        subprocess.run(["git", "-C", str(repo), "config", "user.name", "fixture"], check=True)
        subprocess.run(["git", "-C", str(repo), "config", "user.email", "fixture@invalid"], check=True)
        subprocess.run(["git", "-C", str(repo), "add", "-A"], check=True)
        subprocess.run(["git", "-C", str(repo), "commit", "-q", "-m", "chore: base"], check=True)
        head = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"],
                              capture_output=True, text=True, check=True)
        return head.stdout.strip()

    def test_verify_without_contract_unchanged(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.make_verified_repo(tmp)
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out)

    def test_verify_clean_when_adoption_green(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.make_verified_repo(tmp)
            self.seed_green_tree(repo)
            sha = self.init_git(repo)
            contract = self.green_contract()
            contract["phases"]["6"] = {
                "prs": [{"number": 101, "sha": sha, "conclusion": "success"}]
            }
            (repo / ".github" / "ci-pattern-adoption.json").write_text(
                json.dumps(contract), encoding="utf-8"
            )
            code, out, err = run_cli(
                ["adoption", "generate-doc", str(repo),
                 "--out", str(repo / "docs" / "ci-pattern-flow.md")])
            self.assertEqual(code, 0, err)
            agents = repo / "AGENTS.md"
            doc_text = (repo / "docs" / "ci-pattern-flow.md").read_text(encoding="utf-8")
            begin = "<!-- ci-pattern-agents:begin -->"
            end = "<!-- ci-pattern-agents:end -->"
            agents.write_text(
                agents.read_text(encoding="utf-8")
                + "Flujo operativo: docs/ci-pattern-flow.md\n"
                + doc_text[doc_text.index(begin):doc_text.index(end) + len(end)]
                + "\n", encoding="utf-8")
            self.seal_manifest(repo)
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out)

    def test_verify_fails_when_adoption_incomplete(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.make_verified_repo(tmp)
            self.seed_green_tree(repo)
            self.init_git(repo)
            contract = self.green_contract()
            contract["phases"]["6"] = {"prs": []}
            (repo / ".github" / "ci-pattern-adoption.json").write_text(
                json.dumps(contract), encoding="utf-8"
            )
            self.seal_manifest(repo)
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("fase 5", out + err)

    def test_verify_exit_2_on_invalid_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = self.make_verified_repo(tmp)
            (repo / ".github" / "ci-pattern-adoption.json").write_text("{ no ]", encoding="utf-8")
            code, _, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)


class VerifyFailClosedTests(unittest.TestCase):
    """#182: verify falla cerrado ante manifiesto vacío, P23 mal tipado,
    jobs no estructurales y rutas de manifiesto que escapan del repo."""

    @staticmethod
    def _manifest_with(repo, files, slice_blocks, p23):
        manifest = json.loads((Path(repo) / ".governance-manifest.json").read_text(encoding="utf-8"))
        manifest["files"] = files
        manifest["slice_blocks"] = slice_blocks
        manifest["parameters"]["P23_required_jobs"] = p23
        (Path(repo) / ".governance-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    def test_empty_manifest_rejected(self):
        """Criterio 1: sin nada verificable, exit != 0 con motivo explícito."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={}, slice_blocks={}, p23=[])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertNotEqual(code, 0)
            self.assertEqual(code, 2)
            self.assertIn("nada verificable", out + err)

    def test_p23_string_rejected(self):
        """Criterio 2: 'lint' se iteraba carácter a carácter → limpio."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={}, slice_blocks={}, p23="lint")
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("P23_required_jobs", out + err)

    def test_p23_int_rejected_without_traceback(self):
        """Criterio 2: 5 → TypeError con traceback; ahora exit 2 sin traceback."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={}, slice_blocks={}, p23=5)
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("P23_required_jobs", out + err)
            self.assertNotIn("Traceback", err)

    def test_p23_empty_string_element_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={}, slice_blocks={}, p23=["lint", ""])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("P23_required_jobs", out + err)

    def test_nested_with_key_is_not_a_job(self):
        """Criterio 3: un `lint:` anidado bajo `with:` no es un job cableado."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp, jobs=())
            (repo / ".github" / "workflows" / "ci.yml").write_text(
                "jobs:\n"
                "  build:\n"
                "    runs-on: ubuntu-24.04\n"
                "    steps:\n"
                "      - uses: actions/checkout@v4\n"
                "        with:\n"
                "          lint:\n"
                "            strict: true\n",
                encoding="utf-8",
            )
            self._manifest_with(repo, files={"scripts/check_pr_size.py": sha(repo / "scripts/check_pr_size.py")},
                                slice_blocks={"APAP_WEB": write_slice_block(repo, "APAP_WEB", "contenido del slice\nsegunda línea\n")}, p23=["lint"])
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("gate-not-wired", out)
            self.assertIn("lint", out)

    def test_real_job_under_jobs_is_wired(self):
        """Criterio 3, lado positivo: la clave directa de `jobs:` sí cablea."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, _, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0)

    def test_absolute_manifest_path_rejected(self):
        """Criterio 4: clave absoluta del manifiesto → rechazo, no verificación."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"/etc/hostname": "sha256:" + "0" * 64},
                                slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("ruta", out + err)

    def test_parent_escape_manifest_path_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"../x": "sha256:" + "0" * 64},
                                slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("ruta", out + err)

    def test_windows_drive_manifest_path_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"C:\\x": "sha256:" + "0" * 64},
                                slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("ruta", out + err)

    def test_parameters_null_rejected_without_traceback(self):
        """Red del auditor: parameters null → AttributeError; ahora exit 2."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            manifest = json.loads((Path(repo) / ".governance-manifest.json").read_text(encoding="utf-8"))
            manifest["parameters"] = None
            (Path(repo) / ".governance-manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("parameters", out + err)
            self.assertNotIn("Traceback", err)

    def test_files_list_rejected_without_traceback(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files=["a"], slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("files", out + err)
            self.assertNotIn("Traceback", err)

    def test_slice_blocks_list_rejected_without_traceback(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"scripts/check_pr_size.py": sha(repo / "scripts/check_pr_size.py")},
                                slice_blocks=["APAP_WEB"], p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            self.assertIn("slice_blocks", out + err)
            self.assertNotIn("Traceback", err)

    def test_jobs_four_space_indent_wired(self):
        """Un workflow válido con jobs a 4 espacios: el gate sí está cableado."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp, jobs=())
            (repo / ".github" / "workflows" / "ci.yml").write_text(
                "jobs:\n"
                "    lint:\n"
                "        runs-on: ubuntu-24.04\n",
                encoding="utf-8",
            )
            self._manifest_with(repo, files={"scripts/check_pr_size.py": sha(repo / "scripts/check_pr_size.py")},
                                slice_blocks={"APAP_WEB": write_slice_block(repo, "APAP_WEB", "contenido del slice\nsegunda línea\n")},
                                p23=["lint"])
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0)

    def test_jobs_four_space_nested_with_still_not_wired(self):
        """La indentación flexible no reabre el agujero: un `lint:` anidado
        bajo `with:` a mayor profundidad sigue sin cablear el gate."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp, jobs=())
            (repo / ".github" / "workflows" / "ci.yml").write_text(
                "jobs:\n"
                "    build:\n"
                "        runs-on: ubuntu-24.04\n"
                "        steps:\n"
                "            - uses: actions/checkout@v4\n"
                "              with:\n"
                "                  lint:\n"
                "                      strict: true\n",
                encoding="utf-8",
            )
            self._manifest_with(repo, files={"scripts/check_pr_size.py": sha(repo / "scripts/check_pr_size.py")},
                                slice_blocks={"APAP_WEB": write_slice_block(repo, "APAP_WEB", "contenido del slice\nsegunda línea\n")},
                                p23=["lint"])
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1)
            self.assertIn("gate-not-wired", out)


class VerifyFollowup232Tests(unittest.TestCase):
    """#232: P23 como job id de GitHub, todas las claves inválidas listadas
    antes de leer ningún fichero, y predicado de ruta segura en cli-spec."""

    @staticmethod
    def _manifest_with(repo, files, slice_blocks, p23):
        manifest = json.loads((Path(repo) / ".governance-manifest.json").read_text(encoding="utf-8"))
        manifest["files"] = files
        manifest["slice_blocks"] = slice_blocks
        manifest["parameters"]["P23_required_jobs"] = p23
        (Path(repo) / ".governance-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    def test_p23_invalid_job_id_rejected_naming_entry(self):
        """Un nombre con espacio (o Unicode invisible) no es un job id válido
        de GitHub: exit 2 nombrando la entrada, no un falso gate-not-wired."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={}, slice_blocks={}, p23=["lint", "unit\u200bsuite"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            combined = out + err
            self.assertIn("job id", combined)
            self.assertIn("unit\\u200bsuite", combined)
            self.assertNotIn("Traceback", err)

    def test_p23_valid_github_job_id_passes(self):
        """Job ids con guion bajo, guion y dígitos son válidos y cablean."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp, jobs=("lint_job-1",))
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out + err)

    def test_all_invalid_file_keys_listed_before_any_read(self):
        """Dos claves inválidas → el error lista AMBAS; la validación es
        previa a la lectura de cualquier fichero."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"../escape.txt": "sha256:" + "0" * 64,
                                             "/abs.txt": "sha256:" + "0" * 64},
                                slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 2)
            combined = out + err
            self.assertIn("../escape.txt", combined)
            self.assertIn("/abs.txt", combined)
            self.assertIn("antes de leer", combined)

    def test_normalized_posix_keys_accepted_not_rejected(self):
        """`foo//bar` y `foo/./bar` son seguras (PurePosixPath las normaliza
        y no pueden escapar): producen findings de verificación, no exit 2."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            self._manifest_with(repo, files={"foo//bar": "sha256:" + "0" * 64,
                                             "foo/./baz": "sha256:" + "0" * 64},
                                slice_blocks={}, p23=["lint"])
            code, out, err = run_cli(["verify", str(repo)])
            self.assertEqual(code, 1, out + err)
            self.assertNotIn("ruta relativa segura", out + err)
            # El reporte cita la clave cruda del manifiesto como missing.
            self.assertIn("foo//bar", out)
            self.assertIn("foo/./baz", out)

    def test_cli_spec_documents_safe_path_predicate(self):
        """cli-spec.md define el predicado exacto con ejemplos aceptados y
        rechazados que este suite ejercita."""
        spec = (SKILL_DIR / "references" / "cli-spec.md").read_text(encoding="utf-8")
        self.assertIn("ruta relativa segura", spec)
        self.assertIn("foo//bar", spec)
        self.assertIn("foo/./bar", spec)
        self.assertIn("../", spec)
        self.assertIn("^[A-Za-z_][A-Za-z0-9_-]*$", spec)


class SliceCanonicalResolutionTests(unittest.TestCase):
    """#183: verify --slice-canonical resuelve el cuerpo canónico con la
    misma composición que propagate-team-skills.ps1 (partial.TrimEnd() +
    '\\n\\n' + fragment.TrimEnd()), anclada al marcador @ v<sha7> del
    bloque (sha256 del cuerpo, no el HEAD del catálogo)."""

    @classmethod
    def _make_catalog(cls, canon, registry_consumers="default"):
        """Catálogo con el layout real: partials/ + <slice>/AGENTS.fragment.md
        + fleet/registry.json (registry_consumers=None omite el fichero)."""
        partials = Path(canon) / "partials"
        partials.mkdir(parents=True)
        (partials / "web.md").write_text("# Web\ncontenido partial web\n\n\n", encoding="utf-8")
        (partials / "vba-access.md").write_text("# VBA\ncontenido partial vba\n", encoding="utf-8")
        frag_dir = Path(canon) / "cadete"
        frag_dir.mkdir()
        (frag_dir / "AGENTS.fragment.md").write_text("fragment específico de cadete\n", encoding="utf-8")
        frag_dir2 = Path(canon) / "dysflow"
        frag_dir2.mkdir()
        (frag_dir2 / "AGENTS.fragment.md").write_text("fragment específico de dysflow\n", encoding="utf-8")
        if registry_consumers is not None:
            consumers = cls.REGISTRY["repositories"] if registry_consumers == "default" else registry_consumers
            fleet = Path(canon).parent / "fleet"
            fleet.mkdir()
            (fleet / "registry.json").write_text(
                json.dumps({"schema_version": 1, "repositories": consumers}, indent=2),
                encoding="utf-8",
            )
        return partials, frag_dir

    @staticmethod
    def _compose(partial_text, fragment_text):
        """Réplica de la composición del propagador (TrimEnd + join '\\n\\n')."""
        parts = [t.rstrip() for t in (partial_text, fragment_text) if t is not None]
        return "\n\n".join(parts)

    @staticmethod
    def _install_slice(repo, name, body):
        """Escribe el bloque en AGENTS.md exactamente como el propagador y
        devuelve el hash que el manifiesto debe fijar (sha256 del cuerpo + \\n)."""
        sha7 = hashlib.sha256(body.encode("utf-8")).hexdigest()[:7]
        block = f"<!-- personal-skills:slice:{name} @ v{sha7} -->\n{body}\n<!-- /personal-skills:slice:{name} -->\n"
        agents = Path(repo) / "AGENTS.md"
        agents.write_text(("previo\n\n" if agents.exists() else "") + block, encoding="utf-8")
        return "sha256:" + hashlib.sha256((body + "\n").encode("utf-8")).hexdigest()

    def _manifest_with_slice(self, repo, name, block_hash):
        manifest = json.loads((Path(repo) / ".governance-manifest.json").read_text(encoding="utf-8"))
        manifest["slice_blocks"] = {name: block_hash}
        (Path(repo) / ".governance-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    def test_composed_canonical_matches_exits_zero(self):
        """Criterios 1+2: la composición partial+fragment reproduce el bloque
        del consumer y verify sale 0 sobre el layout real."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            body = self._compose(partial, fragment)
            block_hash = self._install_slice(repo, "CAD_APP", body)
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 0, out)
            self.assertNotIn("canonical-unavailable", out)

    def test_stale_slice_detected_when_catalog_advances(self):
        """Criterio 2: si el catálogo avanzó, el consumer queda desfasado."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(partial, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            # El catálogo avanza (nueva versión del partial) tras propagar.
            (canon / "partials" / "web.md").write_text("# Web\ncontenido partial web V2\n", encoding="utf-8")
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("stale-vs-canonical", out)

    def test_fragment_only_slice_resolves(self):
        """Slice propagado sin partial (fragment-only) también resuelve."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            import shutil
            shutil.rmtree(canon / "partials")
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(None, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 0, out)

    def test_trailing_whitespace_trimmed_like_propagator(self):
        """Criterio 1 byte a byte: TrimEnd del propagador recorta el
        whitespace final del partial; el bloque con salto final cuadra."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")  # termina en saltos extra
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            body = self._compose(partial, fragment)
            block_hash = self._install_slice(repo, "CAD_APP", body)
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 0, out)

    def test_unknown_slice_reports_stale(self):
        """Un slice cuyo marcador no reproduce ninguna composición del
        catálogo (otro consumer o catálogo cambiado) sale como desfasado."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            other_partial = "# Web\nversión que ya no existe en el catálogo\n"
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(other_partial, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("stale-vs-canonical", out)

    REGISTRY = {
        "schema_version": 1,
        "repositories": [
            {"name": "CAD_APP", "canonical_name": "cadete", "primary_type": "web"},
            {"name": "OTHER_APP", "canonical_name": "dysflow", "primary_type": "vba-access"},
        ],
    }

    def test_wrong_consumer_composition_reported(self):
        """Un consumer que lleva la composición válida de OTRO consumer no
        sale limpio: la composición es determinista por registry."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            wrong_partial = (canon / "partials" / "vba-access.md").read_text(encoding="utf-8")
            wrong_fragment = (canon / "dysflow" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(wrong_partial, wrong_fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("stale-vs-canonical", out)

    def test_consumer_absent_from_registry(self):
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon, registry_consumers=[{"name": "OTHER_APP", "canonical_name": "dysflow", "primary_type": "vba-access"}])
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(partial, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("canonical-unavailable", out)
            self.assertIn("registry", out)

    def test_registry_file_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon, registry_consumers=None)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(partial, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("canonical-unavailable", out)
            self.assertIn("fleet/registry.json", out)

    def test_empty_partial_no_false_ambiguity(self):
        """Partial vacío: el propagador lo trata como ausente (truthiness)
        y el cuerpo es fragment-only; no hay ambigüedad ni hallazgo."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            (canon / "partials" / "web.md").write_text("", encoding="utf-8")
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            block_hash = self._install_slice(repo, "CAD_APP", self._compose(None, fragment))
            self._manifest_with_slice(repo, "CAD_APP", block_hash)
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 0, out)

    def test_marker_without_v_reported(self):
        """El marcador debe ser 'v' + sha256(cuerpo)[:7]; sin la 'v' es
        hallazgo explícito, no un limpio silencioso."""
        with tempfile.TemporaryDirectory() as tmp:
            canon = Path(tmp) / "slices"
            self._make_catalog(canon)
            repo, _ = make_repo(tmp, slice_name="CAD_APP")
            partial = (canon / "partials" / "web.md").read_text(encoding="utf-8")
            fragment = (canon / "cadete" / "AGENTS.fragment.md").read_text(encoding="utf-8")
            body = self._compose(partial, fragment)
            sha7 = hashlib.sha256(body.encode("utf-8")).hexdigest()[:7]
            agents = Path(repo) / "AGENTS.md"
            agents.write_text(
                f"<!-- personal-skills:slice:CAD_APP @ {sha7} -->\n{body}\n<!-- /personal-skills:slice:CAD_APP -->\n",
                encoding="utf-8",
            )
            self._manifest_with_slice(repo, "CAD_APP", "sha256:" + hashlib.sha256((body + "\n").encode("utf-8")).hexdigest())
            code, out, _ = run_cli(["verify", str(repo), "--slice-canonical", str(canon)])
            self.assertEqual(code, 1)
            self.assertIn("marker-mismatch", out)


class StatusContractTests(unittest.TestCase):
    """#186: status honra su contrato de salida y verify --json es
    atribuible a una revisión concreta y determinista."""

    def test_status_with_findings_exits_nonzero(self):
        """Criterio 1 (opción A): status con hallazgos sale exit 1, no 0."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            (repo / "scripts" / "check_pr_size.py").write_text("# editado localmente\n")
            code, out, _ = run_cli(["status", str(repo)])
            self.assertEqual(code, 1, out)
            self.assertIn("files", out)

    def test_status_clean_still_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, _, _ = run_cli(["status", str(repo)])
            self.assertEqual(code, 0)

    def test_verify_json_attribution_fields(self):
        """Criterio 2: head_sha, cli_version, schema_version y
        manifest_sha256 presentes y correctos."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            code, out, _ = run_cli(["verify", str(repo), "--json"])
            self.assertEqual(code, 0, out)
            payload = json.loads(out)
            # Fixture sin git: head_sha es None de forma graceful.
            self.assertIsNone(payload["head_sha"])
            # Repo con commit: head_sha se liga al HEAD real.
            subprocess.run(["git", "init", "-q", str(repo)], check=True)
            subprocess.run(["git", "-C", str(repo), "config", "user.email", "t@t"], check=True)
            subprocess.run(["git", "-C", str(repo), "config", "user.name", "t"], check=True)
            subprocess.run(["git", "-C", str(repo), "add", "-A"], check=True)
            subprocess.run(["git", "-C", str(repo), "commit", "-qm", "attribution"], check=True)
            head = subprocess.run(
                ["git", "-C", str(repo), "rev-parse", "HEAD"],
                capture_output=True, text=True, check=True,
            ).stdout.strip()
            code, out, _ = run_cli(["verify", str(repo), "--json"])
            payload = json.loads(out)
            self.assertEqual(code, 0, out)
            self.assertEqual(payload["head_sha"], head)
            self.assertTrue(payload["cli_version"])
            self.assertEqual(payload["schema_version"], 1)
            self.assertEqual(
                payload["manifest_sha256"],
                "sha256:" + hashlib.sha256(
                    (repo / ".governance-manifest.json").read_bytes()
                ).hexdigest(),
            )

    def test_verify_json_deterministic(self):
        """Criterio 3: dos ejecuciones sobre el mismo commit y manifiesto
        producen JSON idéntico byte a byte (nada de timestamps)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            _, out1, _ = run_cli(["verify", str(repo), "--json"])
            _, out2, _ = run_cli(["verify", str(repo), "--json"])
            self.assertEqual(out1, out2)

class ParamsSchemaHardeningTests(unittest.TestCase):
    """#184: los defaults del esquema pasan su propio validador y el
    subconjunto YAML rechaza entradas ambiguas o mal tipadas."""

    @staticmethod
    def _params_file(tmp, extra):
        repo, params = make_repo(tmp)
        params.write_text(VALID_PARAMS + extra, encoding="utf-8")
        return repo, params

    def test_every_schema_default_passes_its_own_validator(self):
        """Criterio 1: barrido de todos los defaults del esquema; falla hoy
        por P19_health_path ('/healthz' con tipo path)."""
        schema, _ = CLI.load_schema()
        bad = []
        for param in schema["parameters"]:
            if "default" not in param or param["default"] is None:
                continue
            err = CLI.validate_value(param, param["default"])
            if err:
                bad.append(err)
        self.assertEqual(bad, [])

    def test_p19_healthz_accepted_as_url_path(self):
        """Criterio 2: P19 usa un tipo de ruta URL; /healthz es válido."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, "P19_health_path: /healthz\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 0, out)

    def test_p19_relative_path_rejected_as_url_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, "P19_health_path: healthz\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P19_health_path", out)

    def test_duplicate_subkey_rejected_with_invalid_input_code(self):
        """Criterio 3: subclaves duplicadas → código de entrada inválida (2)."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(
                tmp, "P22_smoke_retry:\n  attempts: 3\n  attempts: 5\n"
            )
            code, out, err = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 2)
            self.assertIn("subclave duplicada", out + err)

    def test_list_items_validated_against_item_type(self):
        """Criterio 4a: las listas tipan sus elementos (P23 como lista de
        cadenas; un integer dentro es hallazgo)."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, "P23_required_jobs:\n  - lint\n  - 5\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P23_required_jobs", out)

    def test_p48_fields_validated(self):
        """Criterio 4b: el objeto de P48 valida sus campos declarados."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(
                tmp, 'P48_smoke_checks:\n  "/x":\n    status: muchas\n'
            )
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P48_smoke_checks", out)

    def test_p48_unknown_field_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(
                tmp, 'P48_smoke_checks:\n  "/x":\n    foo: 1\n'
            )
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P48_smoke_checks", out)

    def test_p48_valid_contract_accepted(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(
                tmp,
                'P48_smoke_checks:\n'
                '  "/healthz":\n'
                '    status: 200\n'
                '    revision: true\n'
                '  "/":\n'
                '    redirect: /login\n',
            )
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 0, out)

    def test_windows_drive_path_rejected(self):
        """Criterio 5: ruta con letra de unidad de Windows → rechazo."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, 'P41_gate_policy_path: "C:\\tools\\x"\n')
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P41_gate_policy_path", out)

    def test_windows_unc_path_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, 'P41_gate_policy_path: "\\\\server\\share\\x"\n')
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P41_gate_policy_path", out)

    def test_regex_must_be_quoted_is_explicit(self):
        """Criterio 6 (documentación elegida): regex sin comillas no pasa
        silenciosamente como lista; el rechazo nombra el parámetro."""
        with tempfile.TemporaryDirectory() as tmp:
            _, params = self._params_file(tmp, "P01_branch_name_pattern: [a-z]+\n")
            code, out, _ = run_cli(["params", "validate", str(params)])
            self.assertEqual(code, 1)
            self.assertIn("P01_branch_name_pattern", out)


class ReportFrictionTests(unittest.TestCase):
    """#198 slice 1: clasificación explícita, saneado con datos sensibles
    sembrados, huella estable y --dry-run sin red. La publicación y la
    deduplicación llegan en el slice 2."""

    def _friction(self, extra):
        return run_cli([
            "report-friction",
            "--kind", "skill-bug",
            "--gate", "assets/required-jobs/check_required_jobs.py",
            "--skill-version", "0.4", "--skill-sha", "66885f7131fd567fd57cf3ecf1fc366623bcf700",
            "--command", "ci-pattern verify .",
            "--observed", "exit 0 con una violación real",
            "--expected", "exit 1 nombrando la violación",
            "--repro", "1. clonar; 2. ejecutar verify",
            *extra,
        ])

    def test_classification_routes_by_explicit_field(self):
        """La clasificación es el campo --kind, nunca el texto: skill-bug va
        a team-skills; consumer-local se queda en el tracker del consumer."""
        code, out, err = self._friction(["--dry-run", "/tmp/friction-skill.md"])
        self.assertEqual(code, 0, out + err)
        self.assertIn("team-skills", out)
        code, out, err = self._friction(["--kind", "consumer-local",
                                         "--dry-run", "/tmp/friction-local.md"])
        self.assertEqual(code, 0, out + err)
        self.assertIn("consumer", out)
        self.assertNotIn("team-skills", out)

    def test_unknown_kind_fails_closed(self):
        code, _, err = self._friction(["--kind", "maybe", "--dry-run", "/tmp/x.md"])
        self.assertEqual(code, 2)
        self.assertIn("kind", err)

    def test_missing_required_field_fails_closed(self):
        code, _, err = run_cli([
            "report-friction", "--kind", "skill-bug", "--gate", "verify",
            "--skill-version", "0.4", "--skill-sha", "abc",
            "--command", "verify .", "--expected", "exit 1",
            "--dry-run", "/tmp/x.md"])
        self.assertEqual(code, 2)
        self.assertIn("observed", err)

    def test_dry_run_sanitizes_seeded_sensitive_data(self):
        """Saneado: rutas absolutas, usuario, host, token, proyecto privado
        sembrados y un bloque de código del consumer NO llegan al cuerpo."""
        observed = ("Traceback en /home/alonso/proyectopx/app/rutas.py\n"
                    "contacto: alonso@host-interno.corp\n"
                    "token ghp_ABCDEF1234567890abc\n"
                    "```\ncódigo privado del consumer: CLAVE = 'secreto-real'\n```")
        code, out, err = self._friction([
            "--observed", observed,
            "--redact", "proyectopx",
            "--dry-run", "/tmp/friction-sanitized.md"])
        self.assertEqual(code, 0, out + err)
        body = Path("/tmp/friction-sanitized.md").read_text(encoding="utf-8")
        for seeded in ("/home/alonso", "alonso@host-interno.corp", "ghp_ABCDEF1234567890",
                       "proyectopx", "secreto-real", "```"):
            self.assertNotIn(seeded, body)
        self.assertIn("<path>", body)
        self.assertIn("<token>", body)
        self.assertIn("<redacted>", body)

    def test_bare_fqdn_is_sanitized_but_filenames_survive(self):
        """Hallazgo 2: un FQDN suelto (build01.corp.internal) se tapa; un
        nombre de fichero (check_branch_name.py) NO se toca."""
        code, out, err = self._friction([
            "--command", "python3 check_branch_name.py --host build01.corp.internal",
            "--observed", "fallo en build01.corp.internal tras validar check_branch_name.py",
            "--dry-run", "/tmp/friction-fqdn.md"])
        self.assertEqual(code, 0, out + err)
        body = Path("/tmp/friction-fqdn.md").read_text(encoding="utf-8")
        self.assertNotIn("build01.corp.internal", body)
        self.assertIn("check_branch_name.py", body)
        self.assertIn("<host>", body)

    def test_consumer_local_body_speaks_consumer(self):
        """Hallazgo 3: con destino consumer el cuerpo no puede decir que el
        defecto es atribuible a la skill ni mandar a corregirla."""
        code, out, err = self._friction([
            "--kind", "consumer-local", "--dry-run", "/tmp/friction-local.md"])
        self.assertEqual(code, 0, out + err)
        body = Path("/tmp/friction-local.md").read_text(encoding="utf-8")
        self.assertNotIn("atribuible a la skill", body)
        self.assertNotIn("Corregir en la skill", body)
        self.assertIn("consumer", body)

    def test_redact_without_value_is_usage_error_not_traceback(self):
        """Hallazgo 4: --redact sin valor → UsageError exit 2, sin traceback."""
        code, out, err = run_cli(["report-friction", "--redact"])
        self.assertEqual(code, 2)
        self.assertNotIn("Traceback", err)
        self.assertIn("redact", err)


class ReportFrictionPublishTests(unittest.TestCase):
    """#198 slice 2 (+moved tests): deduplicación por huella y publicación
    real, SIEMPRE con `gh` simulado (ningún test toca GitHub ni un
    repositorio real). Resultado verificado leyendo de vuelta:
    confirmed | no_write | unknown. Un solo intento: sin reintentos."""

    def _fake_gh(self, tmp, config):
        """Escribe un `gh` simulado en PATH: registra invocaciones en
        $FAKE_GH_LOG y responde según el config JSON dado."""
        bindir = Path(tmp) / "bin"
        bindir.mkdir(parents=True, exist_ok=True)
        cfg = Path(tmp) / "gh-config.json"
        cfg.write_text(json.dumps(config), encoding="utf-8")
        log = Path(tmp) / "gh-log.txt"
        gh = bindir / "gh"
        gh.write_text(
            "#!/usr/bin/env python3\n"
            "import json, os, sys\n"
            "args = sys.argv[1:]\n"
            "with open(os.environ['FAKE_GH_LOG'], 'a') as log:\n"
            "    log.write(' '.join(args) + '\\n')\n"
            "cfg = json.load(open(os.environ['FAKE_GH_CONFIG']))\n"
            "GH_VIEW_FIELDS = {'assignees', 'author', 'body', 'closed', 'closedAt', 'comments',"
            " 'createdAt', 'id', 'labels', 'milestone', 'number', 'projectCards',"
            " 'projectItems', 'reactionGroups', 'state', 'title', 'updatedAt', 'url'}\n"
            "if '--json' in args:\n"
            "    for field in args[args.index('--json') + 1].split(','):\n"
            "        if field not in GH_VIEW_FIELDS:\n"
            "            print(f'Unknown JSON field: \"{field}\"', file=sys.stderr)\n"
            "            sys.exit(1)\n"
            "if '--search' in args:\n"
            "    cfg['search'] = args[args.index('--search') + 1]\n"
            "    json.dump(cfg, open(os.environ['FAKE_GH_CONFIG'], 'w'))\n"
            "body = cfg['view_body'].replace('aabb', cfg.get('search', ''))\n"
            "if 'issue' in args and 'list' in args:\n"
            "    print(json.dumps(cfg['list']))\n"
            "elif 'issue' in args and 'view' in args:\n"
            "    if cfg.get('fail_view'):\n"
            "        sys.exit(1)\n"
            "    if cfg.get('created'):\n"
            "        print(json.dumps({'number': cfg['create_number'], 'state': cfg['state'].lower(),"
            " 'state_reason': cfg.get('state_reason'), 'body': cfg['created_body']}))\n"
            "    else:\n"
            "        print(json.dumps({'number': cfg['list'][0]['number'] if cfg['list'] else 0,"
            " 'state': cfg['state'].lower(), 'state_reason': cfg.get('state_reason'), 'body': body}))\n"
            "elif 'api' in args and any(a.endswith('/comments') for a in args):\n"
            "    print(json.dumps([{'body': b.replace('aabb', cfg.get('search', ''))}"
            " for b in cfg['comments']]))\n"
            "elif 'api' in args:\n"
            "    n = [a for a in args if '/issues/' in a][0].rsplit('/', 1)[-1]\n"
            "    print(json.dumps({'number': int(n), 'state': cfg['state'].lower(),"
            " 'state_reason': cfg.get('state_reason'),"
            " 'body': body if cfg.get('list') else cfg.get('created_body', '')}))\n"
            "elif 'issue' in args and 'comment' in args:\n"
            "    i = args.index('--body-file')\n"
            "    cfg['comments'].append(open(args[i + 1]).read())\n"
            "    json.dump(cfg, open(os.environ['FAKE_GH_CONFIG'], 'w'))\n"
            "    print('https://github.com/x/issues/80#issuecomment-1')\n"
            "elif 'issue' in args and 'create' in args:\n"
            "    i = args.index('--body-file')\n"
            "    cfg['created_body'] = open(args[i + 1]).read()\n"
            "    cfg['created'] = True\n"
            "    json.dump(cfg, open(os.environ['FAKE_GH_CONFIG'], 'w'))\n"
            "    print(f\"https://github.com/x/issues/{cfg['create_number']}\")\n"
            "elif 'issue' in args and 'edit' in args:\n"
            "    if cfg.get('fail_label'):\n"
            "        sys.exit(1)\n"
            "    cfg['labels_added'] = cfg.get('labels_added', []) + [args[args.index('--add-label') + 1]]\n"
            "    json.dump(cfg, open(os.environ['FAKE_GH_CONFIG'], 'w'))\n"
            "    print('ok')\n"
            "else:\n"
            "    sys.exit(64)\n",
            encoding="utf-8")
        gh.chmod(0o755)
        return {"PATH": str(bindir) + os.pathsep + os.environ["PATH"],
                "FAKE_GH_LOG": str(log), "FAKE_GH_CONFIG": str(cfg)}, log

    FRICTION = ["--kind", "skill-bug",
                "--gate", "assets/required-jobs/check_required_jobs.py",
                "--skill-version", "0.4", "--skill-sha", "66885f7131fd567fd57cf3ecf1fc366623bcf700",
                "--command", "ci-pattern verify .",
                "--observed", "exit 0 con una violación real",
                "--expected", "exit 1 nombrando la violación",
                "--repro", "1. clonar; 2. ejecutar verify"]


    def _publish(self, tmp, env_extra):
        return run_cli(["report-friction", *self.FRICTION,
                        "--publish", "--repo", "upstream/skills"], env_extra=env_extra)


    def test_no_duplicate_creates_issue_and_confirms_by_readback(self):
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {"list": [], "state": "OPEN",
                                           "view_body": "", "comments": [],
                                           "create_number": 91, "fail_view": False})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 0, out + err)
            self.assertIn("result=confirmed", out)
            created = json.loads(Path(env["FAKE_GH_CONFIG"]).read_text()).get("created")
            self.assertTrue(created)
            log_text = Path(log).read_text() if Path(log).exists() else ""
            self.assertEqual(log_text.count("issue create"), 1,
                             "un solo intento: sin reintentos a ciegas")
            self.assertNotIn("status:approved", log_text)


    def test_consumer_local_publish_requires_consumer_repo(self):
        """Hallazgo 1: consumer-local NUNCA publica en team-skills; sin
        --repo del consumer → exit 2, y con --repo team-skills también."""
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {"list": [], "state": "OPEN",
                                           "view_body": "", "comments": [],
                                           "create_number": 91, "fail_view": False})
            friction = ["report-friction", "--kind", "consumer-local",
                        "--gate", "verify", "--skill-version", "0.4",
                        "--skill-sha", "abc", "--command", "verify .",
                        "--observed", "x", "--expected", "y", "--repro", "z",
                        "--publish"]
            code, out, err = run_cli(friction, env_extra=env)
            self.assertEqual(code, 2, out + err)
            self.assertIn("consumer", err)
            self.assertNotIn("issue create", Path(log).read_text() if Path(log).exists() else "")
            code, out, err = run_cli(friction + ["--repo", "DysTelefonica/team-skills"], env_extra=env)
            self.assertEqual(code, 2, out + err)
            self.assertIn("consumer", err)
            self.assertEqual(Path(log).read_text() if Path(log).exists() else "".count("issue create"), 0)


    def test_comments_paginated_and_issue_list_has_limit(self):
        """Hallazgo 2: los comentarios se leen con --paginate y el listado
        de dedup con --limit explícito; con >30 comentarios el recuento es
        correcto (36ª ocurrencia)."""
        marker = "ci-pattern-friction: aabb"
        with tempfile.TemporaryDirectory() as tmp:
            many = [f"Ocurrencia {i}\n<!-- {marker} -->" for i in range(1, 36)]
            env, log = self._fake_gh(tmp, {
                "list": [{"number": 80, "state": "OPEN"}], "state": "OPEN",
                "view_body": f"<!-- {marker} -->", "comments": many,
                "create_number": 91, "fail_view": False})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 0, out + err)
            log_text = Path(log).read_text() if Path(log).exists() else ""
            self.assertIn("--paginate", log_text)
            self.assertIn("--limit", log_text)
            self.assertIn("Ocurrencia 36", out)


    def test_closed_state_reason_drives_recommendation_or_occurrence(self):
        """Hallazgo 3: cerrado NO implica fix publicado. state_reason
        completed → recomienda actualizar (sin escribir); otro motivo
        (not_planned) → comentario de ocurrencia."""
        for reason, expect_write in (("completed", False), ("not_planned", True)):
            with self.subTest(state_reason=reason):
                with tempfile.TemporaryDirectory() as tmp:
                    env, log = self._fake_gh(tmp, {
                        "list": [{"number": 55, "state": "CLOSED"}],
                        "state": "CLOSED", "state_reason": reason,
                        "view_body": "<!-- ci-pattern-friction: aabb -->",
                        "comments": [], "create_number": 91, "fail_view": False})
                    code, out, err = self._publish(tmp, env)
                    self.assertEqual(code, 0, out + err)
                    if expect_write:
                        self.assertIn("Ocurrencia 1", out)
                        self.assertNotIn("recomendación", out)
                    else:
                        self.assertIn("result=no_write", out)
                        self.assertIn("actualizar", out)
                        self.assertNotIn("issue comment", Path(log).read_text() if Path(log).exists() else "")


    def test_closed_duplicate_recommends_update_without_writing(self):
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {
                "list": [{"number": 55, "state": "CLOSED"}],
                "state": "CLOSED", "state_reason": "completed",
                "view_body": "<!-- ci-pattern-friction: aabb --> fixed in 0.5",
                "comments": [], "create_number": 91, "fail_view": False})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 0, out + err)
            self.assertIn("result=no_write", out)
            self.assertIn("actualizar", out)
            self.assertNotIn("issue create", Path(log).read_text() if Path(log).exists() else "")
            self.assertNotIn("issue comment", Path(log).read_text() if Path(log).exists() else "")


    def test_recurrent_label_declared_in_host_contract_and_failure_is_isolated(self):
        """Hallazgo 4: friction:recurrent está declarada host-enforced en
        .github/host-contract.json (HR-34); si gh falla al añadirla, el
        comentario confirmado se separa de la etiqueta fallida (exit 0)."""
        contract = json.loads((SKILL_ROOT.parents[2] / ".github" / "host-contract.json").read_text())
        names = [l["name"] for l in contract["labels"]]
        self.assertIn("friction:recurrent", names)
        self.assertEqual(contract["labels"][names.index("friction:recurrent")]["class"],
                         "host-enforced")
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {
                "list": [{"number": 80, "state": "OPEN"}], "state": "OPEN",
                "view_body": "<!-- ci-pattern-friction: aabb -->",
                "comments": ["Ocurrencia 1\n<!-- ci-pattern-friction: aabb -->"],
                "create_number": 91, "fail_view": False,
                "fail_label": True})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 0, out + err)
            self.assertIn("result=confirmed", out)
            self.assertIn("friction:recurrent fallida", out)
            self.assertIn("result=confirmed", out)


    def test_unknown_when_readback_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {"list": [], "state": "OPEN",
                                           "view_body": "", "comments": [],
                                           "create_number": 91, "fail_view": True})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 1, out + err)
            self.assertIn("result=unknown", out)


    def test_temp_files_are_unpredictable_and_cleaned(self):
        """Hallazgo 5: sin ficheros friction-* predecibles en /tmp; se
        limpian tras publicar."""
        gp = tempfile.gettempdir()
        before = set(p.name for p in Path(gp).glob("friction-*"))
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {"list": [], "state": "OPEN",
                                           "view_body": "", "comments": [],
                                           "create_number": 91, "fail_view": False})
            code, out, err = self._publish(tmp, env)
            self.assertEqual(code, 0, out + err)
        after = set(p.name for p in Path(gp).glob("friction-*"))
        self.assertEqual(before, after)



    def test_dry_run_still_never_touches_gh(self):
        with tempfile.TemporaryDirectory() as tmp:
            env, log = self._fake_gh(tmp, {"list": [], "state": "OPEN",
                                           "view_body": "", "comments": [],
                                           "create_number": 91, "fail_view": False})
            code, out, err = run_cli(["report-friction", *self.FRICTION,
                                      "--dry-run", str(Path(tmp) / "r.md")], env_extra=env)
            self.assertEqual(code, 0, out + err)
            self.assertEqual(Path(log).read_text() if Path(log).exists() else "", "")





class ReportFrictionHintTests(unittest.TestCase):
    """#198 slice 3: los gates que detectan un fallo propio imprimen la
    invocación de report-friction ya rellenada (traceback, entrada no
    soportada, canonical-unavailable)."""

    def test_canonical_unavailable_prints_hint(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            manifest = json.loads((repo / ".governance-manifest.json").read_text())
            manifest["slice_blocks"] = {"APAP_WEB": "sha256:" + "0" * 64}
            (repo / ".governance-manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
            canonical = Path(tmp) / "slices"
            canonical.mkdir()
            code, out, err = run_cli(["verify", str(repo), "--slice-canonical", str(canonical)])
            self.assertEqual(code, 1, out + err)
            self.assertIn("canonical-unavailable", out)
            self.assertIn("report-friction", out + err)


    def test_future_command_prints_hint(self):
        # #276: adopt/update ya no son wave futura (redirigen sin hint);
        # el hint de fricción queda para comandos futuros reales.
        code, out, err = run_cli(["scaffold"])
        self.assertEqual(code, 2)
        self.assertNotIn("report-friction", out + err)
        code, out, err = run_cli(["adopt"])
        self.assertEqual(code, 2)
        self.assertIn("adoption check", out + err)
        self.assertNotIn("report-friction", out + err)


    def test_unexpected_traceback_prints_hint(self):
        original = dict(CLI.COMMANDS)
        CLI.COMMANDS["boom"] = lambda args: (_ for _ in ()).throw(RuntimeError("fallo real del gate"))
        out, err = io.StringIO(), io.StringIO()
        env = dict(os.environ)
        env["CI_PATTERN_SKILL_DIR"] = str(SKILL_DIR)
        old = os.environ.copy()
        os.environ.clear()
        os.environ.update(env)
        try:
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                with self.assertRaises(RuntimeError):
                    CLI.main(["boom"])
        finally:
            os.environ.clear()
            os.environ.update(old)
            CLI.COMMANDS.clear()
            CLI.COMMANDS.update(original)
        self.assertIn("report-friction", out.getvalue() + err.getvalue())





class WindowsPortabilityTests(unittest.TestCase):
    """#187: el CLI produce el mismo veredicto en Linux y Windows y se
    invoca exactamente como se documenta."""

    def test_crlf_copy_of_lf_file_is_not_a_finding(self):
        """Criterio 1: normalización única CRLF→LF para ficheros; una copia
        CRLF de un fichero LF (core.autocrlf) no es hallazgo."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            target = repo / "scripts" / "check_pr_size.py"
            target.write_bytes(target.read_bytes().replace(b"\n", b"\r\n"))
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out)
            self.assertNotIn("locally-modified", out)

    def test_slice_crlf_copy_is_not_a_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            agents = repo / "AGENTS.md"
            agents.write_bytes(agents.read_bytes().replace(b"\n", b"\r\n"))
            code, out, _ = run_cli(["verify", str(repo)])
            self.assertEqual(code, 0, out)

    def test_findings_print_with_cp1252_stdout(self):
        """Criterio 2: con PYTHONIOENCODING=cp1252 el CLI imprime sus
        hallazgos sin UnicodeEncodeError (ejecutado como proceso)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            (repo / "scripts" / "check_pr_size.py").write_text("# editado\n")
            env = dict(os.environ)
            env["PYTHONIOENCODING"] = "cp1252"
            env["CI_PATTERN_SKILL_DIR"] = str(SKILL_DIR)
            proc = subprocess.run(
                [sys.executable, str(CLI_PATH), "verify", str(repo)],
                capture_output=True, text=True, env=env, check=False,
            )
            self.assertEqual(proc.returncode, 1)
            self.assertIn("locally-modified", proc.stdout + proc.stderr)
            self.assertNotIn("UnicodeEncodeError", proc.stderr)

    def test_cli_invokable_as_process(self):
        """Criterio 4: el CLI se invoca como proceso (no como módulo)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, params = make_repo(tmp)
            proc = subprocess.run(
                [sys.executable, str(CLI_PATH), "params", "validate", str(params)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)

    def test_direct_execution_works(self):
        """Criterio 3: ejecución directa documentada (shebang + exec bit)."""
        if sys.platform == "win32":
            self.skipTest("el exec bit es semántica POSIX")
        with tempfile.TemporaryDirectory() as tmp:
            repo, _ = make_repo(tmp)
            proc = subprocess.run(
                [str(CLI_PATH), "verify", str(repo)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)

    def test_exec_bit_in_git_index(self):
        # El modo 100755 es una propiedad del ÍNDICE del catálogo canónico:
        # sólo es observable dentro de un checkout git. En una copia runtime
        # (~/.agents/skills/ci-pattern) o una extracción de git archive no
        # hay repo y el test se salta con motivo explícito (fuera de un repo
        # el bit lo fija el instalador, no git).
        probe = subprocess.run(
            ["git", "-C", str(SKILL_DIR), "rev-parse", "--is-inside-work-tree"],
            capture_output=True, text=True, check=False,
        )
        if probe.returncode != 0 or probe.stdout.strip() != "true":
            self.skipTest(
                "la skill no vive en un checkout git (copia runtime o git archive): "
                "el modo 100755 del índice no es observable aquí; verifíquelo en "
                "el catálogo canónico."
            )
        toplevel = subprocess.run(
            ["git", "-C", str(SKILL_DIR), "rev-parse", "--show-toplevel"],
            capture_output=True, text=True, check=True,
        ).stdout.strip()
        listing = subprocess.run(
            ["git", "-C", toplevel, "ls-files", "-s",
             "personal/ardelperal/ci-pattern/assets/bin/ci-pattern",
             "personal/ardelperal/ci-pattern/assets/branch-name.sh"],
            capture_output=True, text=True, check=True,
        ).stdout
        modes = [line.split()[0] for line in listing.splitlines()]
        self.assertEqual(modes, ["100755", "100755"])

class AssetInventoryTests(unittest.TestCase):
    """#190: asset-inventory.md coincide con el árbol real de la skill y
    su deriva de rutas es detectada automáticamente (HR-10)."""

    DOC = SKILL_DIR / "references" / "asset-inventory.md"
    REPO_ROOT = SKILL_DIR.parents[2]

    def test_no_legacy_skill_prefix(self):
        """El prefijo 'skills/ci-pattern/' ya no existe en el repo: el
        inventario debe citar 'personal/ardelperal/ci-pattern/'."""
        doc = self.DOC.read_text(encoding="utf-8")
        self.assertNotIn("skills/ci-pattern", doc)

    def test_every_cited_skill_path_exists(self):
        doc = self.DOC.read_text(encoding="utf-8")
        cited = sorted(set(re.findall(
            r"`(personal/ardelperal/ci-pattern/[^`\s]+)`", doc
        )))
        self.assertTrue(cited, "el inventario debe citar rutas de la skill")
        missing = [
            p for p in cited
            if not (self.REPO_ROOT / p).exists()
        ]
        self.assertEqual(missing, [])

    def test_governance_wave_assets_are_documented(self):
        """Los activos entregados por la wave de gobernanza figuran en el
        inventario (si no, el mapa de huecos miente)."""
        doc = self.DOC.read_text(encoding="utf-8")
        for rel in (
            "assets/bin/ci-pattern",
            "assets/branch-name.sh",
            "assets/parameters.schema.json",
            "assets/ci-pattern.example.yaml",
            "assets/host-readback/check_host_drift.py",
            "assets/tests/test-ci-pattern-cli.py",
            "references/cli-spec.md",
            "references/delegation-template.md",
        ):
            self.assertIn(rel, doc)


class WaveSiguienteFreeTests(unittest.TestCase):
    """#376: la CLI, su ayuda y su guía no anuncian como próximos los
    subcomandos retirados, y la plantilla canónica se nombra por el artefacto
    que existe. El barrido es explícito —los ficheros que lee quien usa la
    skill— y no incluye este test, que nombra la cadena para poder prohibirla."""

    REPO_ROOT = SKILL_ROOT.parents[2]
    USER_FACING = (
        "SKILL.md",
        "assets/bin/ci-pattern",
        "assets/templates/README.md",
        "references/porting-guide.md",
        "references/cli-spec.md",
    )

    def test_help_does_not_announce_retired_subcommands(self):
        code, out, err = run_cli(["--help"])
        text = out + err
        self.assertEqual(code, 0, text)
        self.assertNotIn("wave siguiente", text)
        self.assertNotIn(
            "wave futura", text,
            "la ayuda no puede prometer una wave: adopt/update están retirados")

    def test_no_user_facing_file_promises_the_retired_wave(self):
        offenders = []
        for rel in self.USER_FACING:
            flat = " ".join((SKILL_ROOT / rel).read_text(encoding="utf-8").split())
            if "wave siguiente" in flat or "wave futura" in flat:
                offenders.append(rel)
        self.assertEqual([], offenders, "ficheros que anuncian la wave retirada")

    def test_templates_readme_names_an_existing_source(self):
        readme = (SKILL_ROOT / "assets" / "templates" / "README.md").read_text(
            encoding="utf-8")
        for rel in (".github/PULL_REQUEST_TEMPLATE.md", ".github/ISSUE_TEMPLATE",
                    "assets/pr-contract/check_pr_contract.py"):
            self.assertIn(rel, readme, f"la plantilla canónica se nombra por '{rel}'")
        self.assertTrue(
            (self.REPO_ROOT / ".github" / "PULL_REQUEST_TEMPLATE.md").is_file(),
            "la fuente que el README nombra tiene que existir en el árbol del patrón")

    def test_skill_does_not_teach_retired_commands(self):
        skill = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
        # El comando retirado, no su vecino `ci-pattern adoption ...`.
        self.assertNotIn("ci-pattern adopt <repo>", skill)
        self.assertNotIn("wave futura", skill)


if __name__ == "__main__":
    unittest.main(verbosity=2)
