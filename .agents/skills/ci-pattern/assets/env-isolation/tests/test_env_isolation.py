#!/usr/bin/env python3
# ci-pattern asset tests — env isolation gate (HR-28, DysTelefonica/team-skills#245)
"""Executable suite for the env isolation gate (black-box, like CI).

HR-28: the suite environment MUST be isolated from the developer's local
``.env``. The gate scans a suite tree and reports:

- any reference to a local ``.env`` file (``source .env``, ``dotenv``, ...)
  — ``.env.example``-style names are not local state and are exempt;
- an environment variable read that is neither assigned locally in the same
  file (shell locals, ``for`` loop variables, arrays) nor declared in the
  versioned allowlist.

Exit codes: 0 clean · 1 findings · 2 invalid input (missing root, missing
or malformed allowlist, empty subject)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_env_isolation.py"
ALLOWLIST = {"env_vars": ["TMPDIR", "HOME", "PATH", "CI"]}


def build(root: Path, files: dict):
    for rel, content in files.items():
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")


def run_gate(root, allowlist, extra=()):
    with tempfile.TemporaryDirectory() as ad:
        a = Path(ad, "allowlist.json")
        a.write_text(json.dumps(allowlist), encoding="utf-8")
        out = subprocess.run([sys.executable, str(GATE), "--root", str(root),
                              "--allowlist", str(a), *extra],
                             capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_suite_using_allowed_vars_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/a.sh": '#!/usr/bin/env bash\n'
                                        'work="$TMPDIR/work"\necho "$HOME"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))

    def test_local_and_loop_vars_are_not_env_reads(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/b.sh": '#!/usr/bin/env bash\n'
                                        'REPO_ROOT="x"\nfor f in a b; do echo "$f$REPO_ROOT"; done\n'
                                        'arr=(one two)\necho "${arr[0]}"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))

    def test_env_example_reference_is_exempt(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/c.sh": '#!/usr/bin/env bash\n'
                                        'echo "copy .env.example to .env.local first"\n'
                                        'cat .env.example\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))

    def test_mapfile_and_read_are_assignments(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/m.sh": '#!/usr/bin/env bash\n'
                                        'mapfile -t SCRIPTS < <(git ls-files)\n'
                                        'read -r FIRST < afile\n'
                                        'read -a PARTS <<< "x y"\n'
                                        'echo "${SCRIPTS[0]}" "${FIRST}" "${PARTS[0]}"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))

    def test_vars_in_comments_are_not_reads(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/c.sh": '#!/usr/bin/env bash\n'
                                        '# uso ${SECRET_TOKEN:-vacio} como fallback\n'
                                        'echo done\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))

    def test_quote_aware_comment_strip(self):
        """R2 auditor #259: un # dentro de comillas es contenido (sh y py); un
        os.environ comentado no es hallazgo."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/q.sh": 'x="a # $SECRET_TOKEN"\necho "$x"\n',
                         "suites/q.py": '# print(os.environ.get("COMMENTED_OUT"))\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("SECRET_TOKEN", out)
            self.assertNotIn("COMMENTED_OUT", out)

    def test_python_env_reads_checked_against_allowlist(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/d.py": 'import os\nprint(os.environ.get("TMPDIR"))\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False))


class Findings(unittest.TestCase):
    def test_dotenv_source_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/a.sh": '#!/usr/bin/env bash\nset -a\nsource .env\nset +a\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("HR-28", out)

    def test_python_dotenv_loader_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/b.py": 'from dotenv import load_dotenv\nload_dotenv()\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("HR-28", out)

    def test_undeclared_env_var_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/c.sh": '#!/usr/bin/env bash\necho "$SECRET_DEPLOY_TOKEN"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("SECRET_DEPLOY_TOKEN", out)

    def test_python_undeclared_getenv_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/d.py": 'import os\nprint(os.getenv("LOCAL_ONLY_VAR"))\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("LOCAL_ONLY_VAR", out)

    def test_findings_cite_file_and_line(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/deep/e.sh": '#!/usr/bin/env bash\n. ./.env\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn("suites/deep/e.sh", out)


class InvalidInput(unittest.TestCase):
    def test_missing_root_exits_two(self):
        code, out = run_gate(Path("/no/such/root"), ALLOWLIST)
        self.assertEqual(code, 2)

    def test_malformed_allowlist_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/a.sh": 'echo hi\n'})
            a = Path(root, "allowlist.json")
            a.write_text("not json", encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--root", str(root),
                                  "--allowlist", str(a)], capture_output=True, text=True)
            self.assertEqual(out.returncode, 2)

    def test_empty_subject_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")  # no suite files at all
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 2)

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


class DomainDeclaration(unittest.TestCase):
    """#315 (HR-28): el gate DECLARA su dominio —qué extensiones escanea— y
    falla en voz alta si el árbol de suites tiene ficheros de un lenguaje
    fuera de ese dominio que el contrato no excluye explícitamente. Un sujeto
    incompleto no es un pase (HR-3)."""

    def test_out_of_domain_source_file_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"a.sh": "#!/bin/sh\necho ok\n",
                         "b.php": "<?php echo getenv('DB_URL');\n"})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn(".php", out)

    def test_the_finding_names_the_declared_domain(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"a.sh": "#!/bin/sh\necho ok\n", "b.ts": "export const x = 1;\n"})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn(".sh", out)
            self.assertIn(".py", out)

    def test_explicitly_excluded_extension_passes(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"a.sh": "#!/bin/sh\necho ok\n", "b.php": "<?php echo 1;\n"})
            code, out = run_gate(root, {**ALLOWLIST, "excluded_extensions": [".php"]})
            self.assertEqual(code, 0, out)

    def test_non_source_files_are_ignored(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"a.sh": "#!/bin/sh\necho ok\n", "notas.md": "# notas\n",
                         "datos.json": "{}\n", ".env.example": "A=1\n"})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 0, out)

    def test_tree_with_only_out_of_domain_files_is_a_finding_not_doubt(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"b.php": "<?php echo 1;\n"})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1)
            self.assertIn(".php", out)

    def test_malformed_excluded_extensions_is_exit_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"a.sh": "#!/bin/sh\necho ok\n"})
            code, _ = run_gate(root, {**ALLOWLIST, "excluded_extensions": "php"})
            self.assertEqual(code, 2)


class Dogfooding(unittest.TestCase):
    """Dogfooding (#195): en team-skills el escaneo real NO puede saltarse —
    la allowlist vive en .github/ y su ausencia es un fallo, no un skip."""

    def is_team_skills(self) -> bool:
        return (ASSET.parents[4] / "testing" / "run-unit-suites.sh").is_file()

    def test_real_tree_is_isolated(self):
        repo_root = ASSET.parents[4]
        allowlist = repo_root / ".github" / "env-isolation-allowlist.json"
        if not self.is_team_skills():
            self.skipTest("instalación fuera de team-skills: sin árbol real que escanear")
        self.assertTrue(allowlist.is_file(),
                        "falta .github/env-isolation-allowlist.json: el dogfooding "
                        "de HR-28 no puede saltarse en team-skills")
        out = subprocess.run([sys.executable, str(GATE),
                              "--root", str(repo_root / "testing"),
                              "--allowlist", str(allowlist)],
                             capture_output=True, text=True)
        self.assertEqual(out.returncode, 0, out.stdout + out.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
