#!/usr/bin/env python3
# ci-pattern asset tests — e2e-evidence gate (DysTelefonica/team-skills#425)
"""Suite del gate de evidencia e2e por PR con cambio de UI.

Cubre los criterios de aceptacion del issue #425:

- UI sin e2e y sin exencion -> rojo (1);
- UI con flujo e2e anadido -> verde (0);
- UI con flujo e2e solo actualizado -> verde (0);
- solo backend (sin UI) -> verde (0);
- exencion sin justificar (seccion vacia) -> rojo (1);
- exencion justificada -> verde (0);
- fail-closed: cuerpo ausente, base inexistente o sujeto no git -> 2;
- parametrizacion: `--ui-glob` y `--non-ui-prefix` reemplazan los
  defaults; el patron `*.spec.{ts,js,py}` matchea las tres extensiones.

Black-box: se ejecuta el gate como subproceso contra repositorios git
temporales. Solo stdlib, sin red. Ejecucion:
`python3 test_e2e_evidence.py`.
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_e2e_evidence.py"


def git(repo: Path, *args: str) -> None:
    proc = subprocess.run(["git", "-C", str(repo), *args],
                          capture_output=True, text=True)
    assert proc.returncode == 0, f"git {args}: {proc.stderr}"


def make_repo(tmp: str, files: dict[str, str]) -> Path:
    repo = Path(tempfile.mkdtemp(prefix="e2e-gate.", dir=tmp))
    git(repo, "init", "-q", "-b", "main")
    git(repo, "config", "user.email", "t@example.invalid")
    git(repo, "config", "user.name", "t")
    for rel, content in files.items():
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", "base")
    # La base queda en main; el trabajo del PR vive en feature, asi
    # `main..HEAD` es exactamente el diff de la PR.
    git(repo, "checkout", "-q", "-b", "feature")
    return repo


def add_commit(repo: Path, files: dict[str, str], msg: str) -> None:
    for rel, content in files.items():
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        if content is None:
            p.unlink()
        elif p.exists():
            p.write_text(p.read_text(encoding="utf-8") + content,
                         encoding="utf-8")
        else:
            p.write_text(content, encoding="utf-8")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", msg)


def run_gate(repo: Path, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(GATE), "--repo", str(repo),
         "--base", "main", *args],
        capture_output=True, text=True)


BASE_FILES = {
    "app/pages/Home.tsx": "export const Home = () => null;\n",
    "tests/e2e_ci/README.md": "flujos e2e del consumer\n",
    "docs/quality-gates.md": "reglas\n",
    "scripts/deploy.sh": "#!/bin/sh\necho ok\n",
}


class E2eEvidenceGateTests(unittest.TestCase):
    """Criterios de aceptacion de #425, black-box sobre git temporal."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory(prefix="e2e-evidence.")
        self.addCleanup(self._tmp.cleanup)
        self.repo = make_repo(self._tmp.name, BASE_FILES)

    def test_ui_without_e2e_is_red(self) -> None:
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        proc = run_gate(self.repo, "--head", "HEAD")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("sin flujo e2e", proc.stderr)

    def test_ui_with_added_e2e_is_green(self) -> None:
        add_commit(self.repo, {
            "app/pages/Home.tsx": "// cambio de UI\n",
            "tests/e2e_ci/home-flow.spec.ts": "test('home', () => {});\n",
        }, "ui + e2e")
        proc = run_gate(self.repo, "--head", "HEAD")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_e2e_pattern_matches_all_declared_extensions(self) -> None:
        # El default `*.spec.{ts,js,py}` tiene que matchear las tres
        # extensiones: fnmatch no expande llamas por si solo.
        for name in ("a.spec.ts", "b.spec.js", "c.spec.py"):
            with self.subTest(name=name):
                add_commit(self.repo, {
                    "app/pages/Home.tsx": f"// ui {name}\n",
                    f"tests/e2e_ci/{name}": f"// flow {name}\n",
                }, f"ui + {name}")
                proc = run_gate(self.repo, "--head", "HEAD")
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_ui_with_updated_e2e_is_green(self) -> None:
        # La invariante admite ACTUALIZAR un flujo: el fichero ya vive en
        # la base y solo se modifica.
        repo = make_repo(self._tmp.name, {
            **BASE_FILES,
            "tests/e2e_ci/home-flow.spec.ts": "test('home', () => {});\n",
        })
        add_commit(repo, {
            "app/pages/Home.tsx": "// cambio de UI\n",
            "tests/e2e_ci/home-flow.spec.ts": "test('home v2', () => {});\n",
        }, "ui + e2e update")
        proc = run_gate(repo, "--head", "HEAD")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_backend_only_is_green_without_e2e(self) -> None:
        add_commit(self.repo, {"docs/quality-gates.md": "otra regla\n"},
                   "docs only")
        proc = run_gate(self.repo, "--head", "HEAD")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("ninguno es UI", proc.stdout)

    def test_exemption_without_justification_is_red(self) -> None:
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        body = Path(self._tmp.name) / "pr-body-empty.md"
        body.write_text("## Tests\n\nalgo\n\n## Exencion e2e\n\n"
                        "## Otra seccion\n\ncosa\n", encoding="utf-8")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--pr-body-file", str(body))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_exemption_justified_is_green(self) -> None:
        add_commit(self.repo, {"app/pages/Home.tsx": "// solo copy\n"},
                   "ui change")
        body = Path(self._tmp.name) / "pr-body-ok.md"
        body.write_text("## Exencion e2e\n\nSolo cambia el texto de un label,"
                        " sin logica ni navegacion nueva.\n", encoding="utf-8")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--pr-body-file", str(body))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_missing_pr_body_fails_closed(self) -> None:
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--pr-body-file", str(Path(self._tmp.name) / "nope.md"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_unresolvable_base_fails_closed(self) -> None:
        proc = run_gate(self.repo, "--head", "HEAD", "--base",
                        "NoSuchRef-425")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_non_git_repo_fails_closed(self) -> None:
        plain = Path(tempfile.mkdtemp(prefix="e2e-nogit.", dir=self._tmp.name))
        proc = subprocess.run(
            [sys.executable, str(GATE), "--repo", str(plain),
             "--base", "main", "--head", "HEAD"],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_custom_ui_glob_replaces_defaults(self) -> None:
        # Con --ui-glob frontend/**, app/** deja de ser superficie de UI:
        # el default se reemplaza, no se acumula.
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--ui-glob", "frontend/**")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_custom_non_ui_prefix_excludes_from_detection(self) -> None:
        # --non-ui-prefix reemplaza la lista blanca default: se declara
        # `docs/**` como UI y se excluye su subdirectorio legacy.
        repo = make_repo(self._tmp.name, {"docs/ui/Panel.tsx": "x\n"})
        add_commit(repo, {"docs/ui/Panel.tsx": "y\n"}, "doc ui change")
        proc = run_gate(repo, "--head", "HEAD",
                        "--ui-glob", "docs/**",
                        "--non-ui-prefix", "docs/ui/")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("ninguno es UI", proc.stdout)

    def test_exemption_section_is_configurable(self) -> None:
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        body = Path(self._tmp.name) / "pr-body-custom.md"
        body.write_text("## Motivo de exclusion visual\n\nNo hay flujo que "
                        "cambiar: el render lo gobierna otro package.\n",
                        encoding="utf-8")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--pr-body-file", str(body),
                        "--exemption-section", "## Motivo de exclusion visual")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_exemption_section_accepts_tilde(self) -> None:
        # Discrepancia 3 del inventario #528 §5: el default del gate es
        # la grafía sin tilde («## Exencion e2e»). Castellano peninsular
        # prescribe «Exención» con tilde; un consumer que documente la
        # exención con la grafía correcta debe pasar el gate con el
        # default del parámetro, sin necesidad de configurar
        # --exemption-section. El matching debe aceptar ambas grafías
        # como equivalentes.
        add_commit(self.repo, {"app/pages/Home.tsx": "// cambio de UI\n"},
                   "ui change")
        body = Path(self._tmp.name) / "pr-body-tilde.md"
        body.write_text("## Exención e2e\n\nSolo cambia el texto de un "
                        "label, sin lógica ni navegación nueva.\n",
                        encoding="utf-8")
        proc = run_gate(self.repo, "--head", "HEAD",
                        "--pr-body-file", str(body))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
