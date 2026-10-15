"""Shared helpers for migrations-immutable tests (#431).

El archivo original tenia los helpers (ASSET, git, make_repo, add_files,
run_gate) duplicados. Aqui se extraen a un modulo compartido para que
las dos partes de los tests (estructura y nuevas) los reutilicen.
"""
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent / "check_migrations_immutable.py"


def git(cwd: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["git", "-C", str(cwd), *args],
        capture_output=True, text=True, check=check,
    )


def make_repo(tmp: Path, files: dict[str, str]) -> Path:
    """Crea un repo git con los ficheros dados en HEAD. tmp es el
    directorio padre; devuelve el path del repo."""
    repo = tmp / "repo"
    repo.mkdir()
    git(repo, "init", "-q", "-b", "main")
    git(repo, "config", "user.email", "test@test")
    git(repo, "config", "user.name", "test")
    for rel, content in files.items():
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", "init")
    return repo


def add_files(repo: Path, files: dict[str, str]) -> None:
    """Anade ficheros al working tree y commitea."""
    for rel, content in files.items():
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", "head")


def run_gate(repo: Path, *args: str) -> subprocess.CompletedProcess:
    """Ejecuta el gate con --no-check-reserved-prefixes.

    Los tests existentes cubren el comportamiento base del gate
    (rename/delete/edit, prefijo mayor, duplicados historicos). El
    check de PRs abiertas contra main requiere un mock de la API y
    se cubre en test_migrations_immutable_reserved.py.
    """
    return subprocess.run(
        [sys.executable, str(ASSET), "--repo", str(repo), "--no-check-reserved-prefixes", *args],
        capture_output=True, text=True, check=False,
    )


def run_gate_full(repo: Path, *args: str, env: dict | None = None) -> subprocess.CompletedProcess:
    """Ejecuta el gate con todos los checks activados (incluido el
    de PRs abiertas). El caller pasa ``env`` con GITHUB_TOKEN y,
    opcionalmente, GITHUB_API_BASE apuntando a un mock local.
    """
    full_env = os.environ.copy()
    if env:
        full_env.update(env)
    return subprocess.run(
        [sys.executable, str(ASSET), "--repo", str(repo), *args],
        capture_output=True, text=True, check=False, env=full_env,
    )


class BaseMigrationTest(unittest.TestCase):
    """Setup comun: tmpdir + cleanup."""

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="mi-431-"))

    def tearDown(self) -> None:
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)
