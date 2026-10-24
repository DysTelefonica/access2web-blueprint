"""El `.github/dependabot.yml` del consumer es el que genera el canon (#1182).

Regla de direccion: un ecosystem por manifiesto, el `directory` de cada uno en
el directorio donde vive ese manifiesto, y el fichero del consumer comparado
con el generado. Solo stdlib, ``unittest`` y fixtures en ``tempfile``: el gate
lee el arbol y el contrato, no la historia.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from config import (  # noqa: E402
    CONSUMER_PATH,
    declared_ecosystems,
    directories_in,
    findings,
    render,
)

BP = {"ecosystems": {"pip": "app", "github-actions": ".", "npm": "e2e"}}


def _repo(tmp: Path, files: dict[str, str]) -> Path:
    repo = tmp / "consumer"
    for rel, content in files.items():
        target = repo / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
    repo.mkdir(parents=True, exist_ok=True)
    return repo


def _bp_files() -> dict[str, str]:
    return {"app/pyproject.toml": "[project]\nname='x'\n",
            ".github/workflows/ci.yml": "name: ci\n",
            "e2e/package.json": "{}\n"}


class TestDeclaration(unittest.TestCase):
    def test_the_declared_ecosystems_are_read(self) -> None:
        self.assertEqual(declared_ecosystems(BP),
                         {"pip": "app", "github-actions": ".", "npm": "e2e"})
        self.assertEqual(
            declared_ecosystems({"ecosystems": {"pip": {"directory": "app"}}}),
            {"pip": "app"})

    def test_a_shape_that_cannot_be_read_gives_nothing(self) -> None:
        self.assertEqual(declared_ecosystems(None), {})
        self.assertEqual(declared_ecosystems({"ecosystems": "app"}), {})
        self.assertEqual(declared_ecosystems({"ecosystems": {"pip": 3}}), {})


class TestRender(unittest.TestCase):
    def test_one_ecosystem_per_declared_manifest(self) -> None:
        text = render(BP)
        for ecosystem in ("pip", "github-actions", "npm"):
            self.assertIn(f"- package-ecosystem: {ecosystem}", text)
        self.assertNotIn("{{ECOSYSTEMS}}", text)
    def test_the_directory_points_at_the_manifest(self) -> None:
        text = render(BP)
        self.assertIn("directory: /app", text)
        self.assertIn("directory: /\n", text)
        self.assertIn("directory: /e2e", text)
        self.assertEqual(directories_in(text),
                         {"pip": "/app", "github-actions": "/", "npm": "/e2e"})
    def test_the_canonical_data_is_there(self) -> None:
        text = render(BP)
        self.assertIn("version: 2", text)
        self.assertIn("open-pull-requests-limit: 5", text)
        self.assertIn("- dependencies", text)
        self.assertIn("- type:chore", text)
        self.assertIn("prefix: chore(platform)", text)
        self.assertIn("prefix: ci(platform)", text)
    def test_the_note_about_the_lock_is_in_the_file(self) -> None:
        """El fichero dice que el lock lo regenera el PR, no Dependabot (#1181)."""
        self.assertIn("lock", render(BP))


class TestFindings(unittest.TestCase):
    def test_a_consumer_that_matches_has_no_findings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {**_bp_files(), CONSUMER_PATH: render(BP)})
            self.assertEqual(findings(repo, BP), [])
    def test_a_missing_consumer_file_is_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = findings(repo, BP)
            self.assertTrue(any(CONSUMER_PATH in item for item in found), found)
    def test_a_directory_without_manifest_is_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {CONSUMER_PATH: render(BP)})
            found = findings(repo, BP)
            self.assertTrue(any("no hay manifiesto" in item for item in found), found)
            self.assertTrue(any("'pip'" in item for item in found), found)
    def test_a_directory_that_points_elsewhere_is_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {
                **_bp_files(),
                CONSUMER_PATH: render(BP).replace("directory: /app", "directory: /")})
            found = findings(repo, BP)
            self.assertTrue(any("'pip'" in item and "/app" in item
                                for item in found), found)
    def test_an_uninvited_ecosystem_is_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {
                **_bp_files(),
                CONSUMER_PATH: render(BP) + ("  - package-ecosystem: docker\n"
                                             "    directory: /docker\n")})
            found = findings(repo, BP)
            self.assertTrue(any("docker" in item for item in found), found)
    def test_a_declaration_without_ecosystems_is_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {CONSUMER_PATH: "version: 2\n"})
            found = findings(repo, {"ecosystems": {}})
            self.assertTrue(found)
            self.assertIn("dependabot", found[0])


if __name__ == "__main__":
    unittest.main()
