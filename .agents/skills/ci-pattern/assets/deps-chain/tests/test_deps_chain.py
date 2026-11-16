"""El manifiesto y su lock cuentan lo mismo (HR-76, #1181).

Cubre la regla de la cadena: un PR que cambia el manifiesto de dependencias y
no cambia su lock es rojo, y se dice cual es el lock que falta y como se
regenera.

Solo stdlib, ``unittest`` y fixtures en ``tempfile``: el gate lee el arbol y la
lista de ficheros cambiados, no la historia de git.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from check_deps_chain import (  # noqa: E402
    ManifestUnreadableError,
    check,
    declares_dependencies,
    locks_of,
    notes,
)

MANIFEST = "[project]\nname = 'x'\ndependencies = ['httpx>=0.27']\n"
TOOL_ONLY = "[tool.mypy]\nstrict = true\n"


def _repo(tmp: Path, files: dict[str, str]) -> Path:
    repo = tmp / "consumer"
    for rel, content in files.items():
        target = repo / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
    repo.mkdir(parents=True, exist_ok=True)
    return repo


def _bp_files() -> dict[str, str]:
    """El layout de BP: manifiesto en app/, lock hasheado al lado y uv.lock."""
    return {"app/pyproject.toml": MANIFEST,
            "app/uv.lock": "version = 1\n",
            "app/requirements-dev.lock": "httpx==0.27.0 --hash=sha256:x\n",
            "pyproject.toml": TOOL_ONLY}


class TestManifest(unittest.TestCase):
    def test_a_manifest_that_declares_dependencies_is_one(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            self.assertTrue(declares_dependencies(repo / "app" / "pyproject.toml"))

    def test_a_tool_only_manifest_is_not_a_manifest_of_dependencies(self) -> None:
        """El `pyproject.toml` raiz de un consumer solo lleva configuracion."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            self.assertFalse(declares_dependencies(repo / "pyproject.toml"))

    def test_the_locks_are_the_ones_next_to_the_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            locks = [lock.name for lock in locks_of(repo / "app" / "pyproject.toml")]
            self.assertEqual(locks, ["uv.lock", "requirements-dev.lock"])


class TestCheck(unittest.TestCase):
    def test_a_manifest_without_its_lock_is_a_finding(self) -> None:
        """C1: el manifiesto cambio y el lock no -> rojo, nombrando el lock."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = check(repo, ["app/pyproject.toml"])
            self.assertEqual(len(found), 1, found)
            self.assertIn("app/pyproject.toml", found[0])
            self.assertIn("app/uv.lock", found[0])
            self.assertIn("uv lock --project app", found[0])

    def test_a_manifest_with_its_lock_has_no_finding(self) -> None:
        """C2: manifiesto y lock juntos -> verde."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = check(repo, ["app/pyproject.toml", "app/uv.lock",
                                 "app/requirements-dev.lock"])
            self.assertEqual(found, [])

    def test_a_lock_that_changes_alone_is_not_a_finding(self) -> None:
        """Cerrar la cadena sin tocar el manifiesto no es un cambio de dependencias."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            self.assertEqual(check(repo, ["app/uv.lock"]), [])

    def test_a_tool_only_manifest_is_not_measured(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            self.assertEqual(check(repo, ["pyproject.toml"]), [])

    def test_a_manifest_without_a_lock_is_a_note_and_not_a_finding(self) -> None:
        """C3: sin lock declarado no hay coherencia que exigir."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"app/pyproject.toml": MANIFEST})
            self.assertEqual(check(repo, ["app/pyproject.toml"]), [])
            self.assertTrue(notes(repo, ["app/pyproject.toml"]))

    def test_an_unreadable_manifest_fails_closed(self) -> None:
        """C4: un manifiesto ilegible es rojo, no verde."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"app/pyproject.toml": "[project\n"})
            with self.assertRaises(ManifestUnreadableError):
                declares_dependencies(repo / "app" / "pyproject.toml")


class TestEveryLockMustMove(unittest.TestCase):
    """#1693: cada lock del manifiesto tiene que cambiar; uno solo no cierra la cadena.

    El gate miraba si CUALQUIER lock cambiaba: un PR que regenera `uv.lock` dejaba pasar
    un `requirements-dev.lock` desfasado, que es el que instala la CI."""

    def test_a_manifest_with_only_one_of_its_locks_is_a_finding(self) -> None:
        """RED: manifiesto + uv.lock, sin requirements-dev.lock -> rojo que nombra el que falta."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = check(repo, ["app/pyproject.toml", "app/uv.lock"])
            self.assertEqual(len(found), 1, found)
            self.assertIn("app/requirements-dev.lock", found[0])
            self.assertNotIn("app/uv.lock,", found[0])

    def test_the_finding_names_the_missing_lock_and_its_regeneration_command(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = check(repo, ["app/pyproject.toml", "app/uv.lock"])
            self.assertIn("uv export --project app", found[0])
            self.assertIn("--output-file app/requirements-dev.lock", found[0])

    def test_the_other_direction_is_also_a_finding(self) -> None:
        """manifiesto + requirements-dev.lock, sin uv.lock -> rojo que nombra uv.lock."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            found = check(repo, ["app/pyproject.toml", "app/requirements-dev.lock"])
            self.assertEqual(len(found), 1, found)
            self.assertIn("uv lock --project app", found[0])

    def test_every_lock_changed_with_the_manifest_is_green(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), _bp_files())
            self.assertEqual(check(repo, ["app/pyproject.toml", "app/uv.lock",
                                          "app/requirements-dev.lock"]), [])

    def test_every_missing_lock_is_named_when_several_are_missing(self) -> None:
        files = _bp_files()
        files["app/poetry.lock"] = "[[package]]\n"
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), files)
            found = check(repo, ["app/pyproject.toml", "app/uv.lock"])
            self.assertEqual(len(found), 1, found)
            self.assertIn("app/requirements-dev.lock", found[0])
            self.assertIn("app/poetry.lock", found[0])
            self.assertNotIn("(app/uv.lock", found[0])


if __name__ == "__main__":
    unittest.main()
