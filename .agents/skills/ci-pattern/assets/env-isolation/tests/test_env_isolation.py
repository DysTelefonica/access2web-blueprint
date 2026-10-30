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
import importlib.util
import inspect
import json
import subprocess
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_env_isolation.py"
SNIPPET = ASSET / "snippets" / "pytest-socket" / "conftest.py"
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

    def test_dotenv_in_python_docstring_is_not_a_finding(self):
        """Refs #618: DOTENV must not match literal '.env' inside Python
        triple-quoted docstrings. Prosa is not a runtime reference to the
        developer's .env file."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/doc.py":
                         '"""Module docstring mentioning ``.env`` as runtime example."""\n'
                         'def f():\n'
                         '    """The fixture wipes the developer\'s ``.env`` before the body."""\n'
                         '    return None\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_python_write_only_environ_assign_is_not_a_read(self):
        """Refs #618: PY_READ must not match `os.environ["X"] = value` as a
        read. The LHS of `=` is a write; only RHS references are reads."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/write.py":
                         'import os\n'
                         'os.environ["PYTEST_CANARY"] = "value"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_python_environ_ref_inside_docstring_is_not_a_read(self):
        """Refs #618: PY_READ must not match `os.environ['X']` references
        that appear inside Python triple-quoted docstrings. Documentation
        of an API is not a real environment read at test time."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/doc_env.py":
                         '"""os.environ[\'X\'] is mentioned in the docstring."""\n'
                         'def f():\n'
                         '    """Use os.environ.get(\'X\') to read X."""\n'
                         '    return None\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual((code, "HR-28" in out), (0, False), out)


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

    def test_python_environ_rhs_read_still_finds(self):
        """Refs #618 (regression): `x = os.environ[\"X\"]` (RHS) is a real read
        and must keep being flagged. The fix for LHS-as-write must not silence
        RHS reads."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/rhs.py":
                         'import os\n'
                         'x = os.environ["UNLISTED_VAR"]\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1, out)
            self.assertIn("UNLISTED_VAR", out)

    def test_python_environ_equality_still_finds(self):
        """Refs #618 (regression): `os.environ[\"X\"] == \"y\"` is a comparison
        that reads the env. Must keep being flagged."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/eq.py":
                         'import os\n'
                         'if os.environ["UNLISTED_VAR"] == "x":\n'
                         '    pass\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1, out)
            self.assertIn("UNLISTED_VAR", out)

    def test_python_environ_aug_assign_still_finds(self):
        """Refs #618 (regression): `os.environ[\"X\"] += \"y\"` reads then
        writes. Must keep being flagged as a read."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/aug.py":
                         'import os\n'
                         'os.environ["UNLISTED_VAR"] += "x"\n'})
            code, out = run_gate(root, ALLOWLIST)
            self.assertEqual(code, 1, out)
            self.assertIn("UNLISTED_VAR", out)


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


class TestDomainDeclaration(unittest.TestCase):
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


class TestDotenvPrecision(unittest.TestCase):
    """#635: DOTENV casaba `templates.env` (acceso a atributo) y `dotenv` como
    palabra suelta. Solo el fichero `.env`, el cargador y las lecturas de
    entorno son hallazgos; los datos declarados por fichero no."""

    def scan(self, files, allowlist=ALLOWLIST):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, files)
            return run_gate(root, allowlist)

    def test_jinja_env_attribute_is_not_a_dotenv_reference(self):
        code, out = self.scan({"suites/t.py":
                               'def f(templates):\n    return templates.env.from_string("x")\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_uppercase_env_attribute_is_not_a_dotenv_reference(self):
        code, out = self.scan({"suites/t.py":
                               'TEMPLATES = object()\nautoescape = TEMPLATES.env.autoescape\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_dotenv_word_in_url_is_not_a_loader(self):
        code, out = self.scan({"suites/t.py":
                               'URL = "postgresql://dotenv"\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_open_dotenv_file_is_a_finding(self):
        code, out = self.scan({"suites/o.py": 'open(".env").read()\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28", out)

    def test_dotenv_file_in_fixture_is_a_finding_without_declaration(self):
        code, out = self.scan({"suites/f.py":
                               'from pathlib import Path\n'
                               'def t():\n'
                               '    Path(".env").write_text("X=1")\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28", out)

    def test_declared_data_file_may_mention_dotenv(self):
        allow = {**ALLOWLIST, "dotenv_data_files": [
            {"file": "suites/f.py", "reason": "fixture escribe un .env propio"}]}
        code, out = self.scan({"suites/f.py":
                               'from pathlib import Path\n'
                               'def t():\n'
                               '    Path(".env").write_text("X=1")\n'}, allow)
        self.assertEqual(code, 0, out)

    def test_declaration_does_not_exempt_a_loader_in_the_same_file(self):
        allow = {**ALLOWLIST, "dotenv_data_files": [
            {"file": "suites/f.py", "reason": "fixture escribe un .env en tmp_path"}]}
        code, out = self.scan({"suites/f.py":
                               'from dotenv import load_dotenv\n'
                               'load_dotenv()\n'}, allow)
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28", out)

    def test_declaration_without_reason_is_exit_two(self):
        allow = {**ALLOWLIST, "dotenv_data_files": [{"file": "suites/f.py", "reason": " "}]}
        code, out = self.scan({"suites/f.py": 'x = 1\n'}, allow)
        self.assertEqual(code, 2, out)

    def test_stale_declaration_is_a_finding(self):
        allow = {**ALLOWLIST, "dotenv_data_files": [
            {"file": "suites/f.py", "reason": "ya no escribe .env"}]}
        code, out = self.scan({"suites/f.py": 'x = 1\n'}, allow)
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28 STALE", out)

    def test_b_dotenv_built_under_tmp_path_is_not_a_finding(self):
        code, out = self.scan({"suites/t.py":
                               'def t(tmp_path):\n'
                               '    env_file = tmp_path / ".env"\n'
                               '    (tmp_path / ".env").write_text("X=1")\n'
                               '    assert Path(tmp_path) / ".env" != env_file.parent\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_b_dotenv_outside_tmp_path_is_still_a_finding(self):
        code, out = self.scan({"suites/t.py":
                               'def t(tmp_path):\n'
                               '    (tmp_path / "sub").mkdir()\n'
                               '    return open(".env").read()\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("t.py:3", out)

    def test_c_gitignore_assertions_are_not_findings(self):
        code, out = self.scan({"suites/test_ignore.py":
                               'import re, subprocess\n'
                               'def t():\n'
                               '    assert _check_ignored(".env")\n'
                               '    assert ok, ".env must stay ignored by git"\n'
                               '    pat = re.compile(r"^\\s*(?:\\*\\*[/\\\\]\\.env|x)")\n'
                               '    subprocess.run(["git", "log", "--", ".env"])\n'
                               '    subprocess.run(["git", "check-ignore", ".env"])\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_c_reading_dotenv_next_to_an_assertion_is_still_a_finding(self):
        code, out = self.scan({"suites/t.py":
                               'def t():\n'
                               '    assert Path(".env").read_text()\n'
                               '    assert open(".env")\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("t.py:2", out)
        self.assertIn("t.py:3", out)

    def test_c_git_show_of_dotenv_content_is_still_a_finding(self):
        code, out = self.scan({"suites/t.py":
                               'import subprocess\n'
                               'subprocess.run(["git", "show", "HEAD:.env"])\n'})
        self.assertEqual(code, 1, out)

    def test_shell_source_dotenv_is_still_a_finding(self):
        code, out = self.scan({"suites/s.sh": "#!/usr/bin/env bash\nsource .env\n"})
        self.assertEqual(code, 1, out)

    def test_jinja_env_get_template_attribute_is_not_a_finding(self):
        """Refs #851: `TEMPLATES.env.get_template(template_name).render(...)`
        es acceso al binding local `TEMPLATES` (entorno Jinja2), NO una
        referencia al `.env` del desarrollador. Caso verbatim de
        `ardelperal/APAP_WEB:tests/test_xss_audit.py:101`."""
        code, out = self.scan({"suites/t.py":
                               'def render(templates, template_name, context):\n'
                               '    TEMPLATES = templates\n'
                               '    merged = {"base_template": "base.html", **context}\n'
                               '    return TEMPLATES.env.get_template(template_name).render(**merged)\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_dotenv_substring_in_schema_name_is_not_a_finding(self):
        r"""Refs #851: ``dotenv`` como substring de un identificador o nombre
        de schema (``"dotenv_schema"``, ``"postgresql://dotenv"``) es un dato,
        no un cargador. El detector matchea ``\.env\b``, no la palabra
        ``dotenv``. Caso verbatim de
        ``ardelperal/APAP_WEB:tests/test_verify_fallback_ready_cli.py:295``."""
        code, out = self.scan({"suites/t.py":
                               'def test_falls_back_cleanly():\n'
                               '    assert events == [\n'
                               '        ("provision", "postgresql://dotenv"),\n'
                               '        ("drop", ("postgresql://dotenv", "dotenv_schema")),\n'
                               '    ]\n'})
        self.assertEqual((code, "HR-28" in out), (0, False), out)

    def test_dotenv_values_with_dotenv_filename_is_a_finding(self):
        """Refs #851: `dotenv_values('.env')` lee el `.env` del
        desarrollador. La coincidencia viene del literal `.env`
        (DOTENV_FILE), no del cargador (DOTENV_LOADER solo conoce
        `load_dotenv`). Cierra el contrato positivo para el cuarto caso
        explícito que la suite ya tenia cubierto parcialmente con
        `open('.env')` y `Path('.env')`."""
        code, out = self.scan({"suites/t.py":
                               'from dotenv import dotenv_values\n'
                               'def t():\n'
                               '    return dotenv_values(".env")\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28", out)

    def test_load_dotenv_call_is_a_finding(self):
        """Refs #851: una llamada `load_dotenv()` sola (sin `from dotenv
        import`) sigue siendo hallazgo via la primera alternativa de
        DOTENV_LOADER (`\bload_dotenv\b`). Cierra el contrato positivo
        cuando solo aparece la llamada."""
        code, out = self.scan({"suites/t.py":
                               'def t():\n'
                               '    load_dotenv()\n'})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28", out)


class TestOverAllowlist(unittest.TestCase):
    """#1030: la dirección contraria —una entrada de ``env_vars`` que la suite ya
    no lee. El gate ya la calculaba para ``dotenv_data_files`` (#635) y no para
    ``env_vars``: una entrada de más pasaba en verde sin que nada lo advirtiese.

    El control sale detrás de ``--strict-over-allowlist`` (radio de impacto: la
    allowlist de cada consumer lleva meses escribiendo y ninguna poda sería
    revisada el día que el control entre sin pedirlo)."""

    OK_LINE = ("ENV-ISOLATION OK: {n} suite file(s) scanned — no .env reads, "
               "no undeclared environment variables.")

    # Lee las cuatro entradas de ALLOWLIST con clave literal: sin esta lectura
    # literal, cada una sería además candidata a STALE y el caso no distinguiría
    # la entrada sobrante de las que sí se leen.
    READS_ALL = ('import os\n'
                 'print(os.environ["TMPDIR"], os.environ["HOME"],\n'
                 '      os.environ["PATH"], os.environ["CI"])\n')

    def scan(self, files, allowlist=ALLOWLIST, strict=True):
        extra = ("--strict-over-allowlist",) if strict else ()
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, files)
            return run_gate(root, allowlist, extra)

    def allow_with(self, *extra_vars):
        return {**ALLOWLIST, "env_vars": ALLOWLIST["env_vars"] + list(extra_vars)}

    def test_unused_entry_is_a_stale_finding_naming_the_variable(self):
        code, out = self.scan({"suites/a.py": self.READS_ALL},
                              self.allow_with("GITHUB_TOKEN"))
        self.assertEqual(code, 1, out)
        self.assertIn("HR-28 STALE", out)
        self.assertIn("GITHUB_TOKEN", out)

    def test_stale_reads_differently_from_an_undeclared_read(self):
        """Un operador tiene que poder distinguir «sobra una entrada» de
        «falta declararla»: son arreglos opuestos en el mismo fichero."""
        _, stale = self.scan({"suites/a.py": self.READS_ALL},
                             self.allow_with("GITHUB_TOKEN"))
        _, undeclared = self.scan({"suites/a.py":
                                   self.READS_ALL + 'print(os.environ["RUNNER_SECRET"])\n'})
        self.assertIn("not declared in the versioned allowlist", undeclared)
        stale_only = [ln for ln in stale.splitlines() if "STALE" in ln]
        self.assertEqual(len(stale_only), 1, stale)
        self.assertNotIn("not declared in the versioned allowlist", stale_only[0])

    def test_unused_is_decided_after_the_whole_tree_is_scanned(self):
        """La dirección contraria no puede decidirse con el estado parcial de un
        fichero: una entrada leída sólo en el último fichero alfabético, y no en
        el primero, no sale como STALE; otra que no se lee en ninguno, sí."""
        code, out = self.scan({
            "suites/a.py": 'import os\nprint(os.environ["TMPDIR"], os.environ["HOME"],\n'
                           '      os.environ["PATH"])\n',
            "suites/z.py": 'import os\nprint(os.environ["CI"], os.environ["LATE_READ"])\n'},
            self.allow_with("LATE_READ", "GITHUB_TOKEN"))
        self.assertEqual(code, 1, out)
        self.assertEqual(out.count("HR-28 STALE"), 1, out)
        self.assertIn("GITHUB_TOKEN", out)
        self.assertNotIn("LATE_READ", out)

    def test_entry_read_only_from_a_shell_file_is_not_a_finding(self):
        """El barrido es mixto: SH_READ cubre la lectura de shell, así que una
        entrada leída sólo desde un ``*.sh`` no es sobre-allowlisting."""
        code, out = self.scan({"suites/a.sh": '#!/usr/bin/env bash\n'
                                              'work="$FEATURE_WORK_DIR"\n'
                                              'echo "$work" "$TMPDIR" "$HOME" "$PATH" "$CI"\n'},
                              {**ALLOWLIST, "env_vars": ALLOWLIST["env_vars"]
                               + ["FEATURE_WORK_DIR"]})
        self.assertEqual(code, 0, out)
        self.assertNotIn("FEATURE_WORK_DIR", out)

    def test_bulk_read_is_not_a_false_positive(self):
        """``os.environ.copy()`` no se puede atribuir a un nombre: la entrada sin
        lectura literal no se afirma sobrante, se informa como no probable y no
        bloquea."""
        code, out = self.scan({"suites/a.py": self.READS_ALL
                               + 'import subprocess\n'
                                 'env = os.environ.copy()\n'
                                 'subprocess.run(["echo"], env=env)\n'},
                              self.allow_with("GITHUB_TOKEN"))
        self.assertEqual(code, 0, out)
        self.assertNotIn("HR-28 STALE", out)
        self.assertIn("HR-28 BLOCK-READ", out)
        self.assertIn("GITHUB_TOKEN", out)

    def test_every_bulk_read_form_is_recognised(self):
        forms = ("env = os.environ.copy()\n",
                 "for k, v in os.environ.items():\n    print(k, v)\n",
                 "print(sorted(os.environ.keys()))\n",
                 "env = dict(os.environ)\n",
                 "NAME = 'GITHUB_TOKEN'\nos.getenv(NAME)\n",
                 "print(os.environ[NAME])\n")
        for form in forms:
            with self.subTest(form=form):
                code, out = self.scan({"suites/a.py": "import os\n" + self.READS_ALL + form},
                                      self.allow_with("GITHUB_TOKEN"))
                self.assertEqual(code, 0, out)
                self.assertIn("HR-28 BLOCK-READ", out)

    def test_bulk_read_does_not_silence_an_undeclared_read(self):
        """No-regresión: la lectura sin declarar sigue siendo hallazgo aunque el
        árbol lea el entorno en bloque."""
        code, out = self.scan({"suites/a.py": "import os\n" + self.READS_ALL
                               + 'env = os.environ.copy()\n'
                                 'print(os.environ["RUNNER_SECRET"])\n'},
                              self.allow_with("GITHUB_TOKEN"))
        self.assertEqual(code, 1, out)
        self.assertIn("RUNNER_SECRET", out)
        self.assertIn("not declared in the versioned allowlist", out)

    def test_a_write_to_the_environment_is_not_a_read(self):
        """Escribir en el entorno no prueba que una entrada siga usándose: el
        contrato de la allowlist es lo que las suites PUEDEN LEER."""
        code, out = self.scan({"suites/a.py":
                               'import os\n'
                               'print(os.environ["TMPDIR"], os.environ["HOME"],\n'
                               '      os.environ["PATH"])\n'
                               'os.environ["FEATURE_FLAG"] = "1"\n'},
                              {"env_vars": ["TMPDIR", "HOME", "PATH", "FEATURE_FLAG"]})
        self.assertEqual(code, 1, out)
        self.assertEqual(out.count("HR-28 STALE"), 1, out)
        self.assertIn("'FEATURE_FLAG'", out)

    def test_a_shell_local_is_not_a_read_of_the_environment(self):
        """Un valor que el propio fichero asigna es local (la regla que el gate
        ya aplica a la dirección de lectura-no-declarada, HR-1)."""
        code, out = self.scan({"suites/a.sh": '#!/usr/bin/env bash\n'
                                              'CI="1"\n'
                                              'echo "$CI" "$TMPDIR" "$HOME" "$PATH"\n'})
        self.assertEqual(code, 1, out)
        self.assertEqual(out.count("HR-28 STALE"), 1, out)
        self.assertIn("'CI'", out)

    def test_the_flag_is_opt_in_so_the_day_it_lands_breaks_nobody(self):
        """Sin la bandera, una entrada de más no falla: el control se calcula y
        se nombra, y sale 1 sólo cuando se pide."""
        code, out = self.scan({"suites/a.py": self.READS_ALL},
                              self.allow_with("GITHUB_TOKEN"), strict=False)
        self.assertEqual(code, 0, out)
        self.assertIn(self.OK_LINE.format(n=1), out)
        self.assertNotIn("HR-28", out)

    def test_clean_allowlist_keeps_exit_zero_and_the_ok_message(self):
        code, out = self.scan({"suites/a.py": self.READS_ALL}, ALLOWLIST)
        self.assertEqual(code, 0, out)
        self.assertIn(self.OK_LINE.format(n=1), out)

    def test_the_issue_fixture_pins_both_directions_at_once(self):
        """Los cuatro casos en un solo árbol, porque su interacción es lo que hay
        que fijar: leída con clave literal, leída con clave no literal, declarada
        y no leída, y leída desde shell."""
        allow = {**ALLOWLIST, "env_vars": ALLOWLIST["env_vars"]
                 + ["LITERAL_READ", "NON_LITERAL_READ", "DECLARED_NEVER_READ",
                    "SHELL_ONLY_READ"]}
        code, out = self.scan({
            "suites/a.py": 'import os\n'
                           'print(os.environ["TMPDIR"], os.environ["HOME"],\n'
                           '      os.environ["PATH"], os.environ["CI"],\n'
                           '      os.environ["LITERAL_READ"])\n'
                           'NAME = "NON_LITERAL_READ"\n'
                           'print(os.getenv(NAME))\n',
            "suites/z.sh": '#!/usr/bin/env bash\necho "$SHELL_ONLY_READ"\n',
        }, allow)
        self.assertEqual(code, 0, out)
        self.assertIn(self.OK_LINE.format(n=2), out)
        self.assertIn("DECLARED_NEVER_READ", out)   # la entrada que sobra
        self.assertIn("NON_LITERAL_READ", out)      # la limitación conocida
        self.assertNotIn("LITERAL_READ'", out.split("BLOCK-READ")[1])
        self.assertNotIn("SHELL_ONLY_READ'", out.split("BLOCK-READ")[1])

    def test_invalid_input_stays_fail_closed_under_the_flag(self):
        """HR-3 intacta: la bandera no abre una puerta trasera al exit 2."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d, "testing")
            build(root, {"suites/a.sh": "echo hi\n"})
            a = Path(root, "allowlist.json")
            a.write_text("not json", encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--root", str(root),
                                  "--allowlist", str(a), "--strict-over-allowlist"],
                                 capture_output=True, text=True)
            self.assertEqual(out.returncode, 2)
        code, out = self.scan({"suites/a.py": self.READS_ALL}, ALLOWLIST)
        self.assertEqual(code, 0, out)

    def test_the_module_docstring_names_the_second_direction_and_its_exception(self):
        """El docstring del módulo es la documentación que se propaga al consumer:
        tiene que nombrar la dirección que este control vigila y su excepción."""
        spec = importlib.util.spec_from_file_location("gate_doc", GATE)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        doc = module.__doc__
        self.assertIn("no read", doc)
        self.assertIn("--strict-over-allowlist", doc)
        self.assertIn("os.environ.copy()", doc)


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


class TestPytestSocketSnippet(unittest.TestCase):
    """The ``allow_network`` fixture must close the network again at teardown (#663)."""

    def load(self, calls):
        """Import the snippet against stub ``pytest`` / ``pytest_socket`` modules."""
        fake_socket = types.ModuleType("pytest_socket")
        fake_socket.enable_socket = lambda allowed_hosts=None: calls.append(("enable", allowed_hosts))
        fake_socket.disable_socket = lambda: calls.append(("disable", None))
        fake_pytest = types.ModuleType("pytest")
        fake_pytest.fixture = lambda func: func
        fake_pytest.skip = lambda reason: calls.append(("skip", reason))
        with mock.patch.dict(sys.modules, {"pytest": fake_pytest, "pytest_socket": fake_socket}):
            spec = importlib.util.spec_from_file_location("snippet_conftest", SNIPPET)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
        return module

    def test_fixture_yields_the_handle_and_restores_after_the_test(self):
        calls = []
        module = self.load(calls)
        self.assertTrue(inspect.isgeneratorfunction(module.allow_network))
        fixture = module.allow_network()
        handle = next(fixture)
        handle(["api.example.com", "db.example.com"])
        self.assertEqual(calls, [("enable", "api.example.com,db.example.com")])
        with self.assertRaises(StopIteration):
            next(fixture)
        self.assertEqual(calls[-1], ("disable", None))

    def test_restore_is_not_called_before_the_test_finishes(self):
        calls = []
        fixture = self.load(calls).allow_network()
        next(fixture)
        self.assertNotIn(("disable", None), calls)


if __name__ == "__main__":
    unittest.main(verbosity=2)
