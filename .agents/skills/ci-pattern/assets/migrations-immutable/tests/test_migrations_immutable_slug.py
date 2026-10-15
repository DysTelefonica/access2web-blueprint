#!/usr/bin/env python3
# ci-pattern asset tests — migrations-immutable, parte 4: parse_repo_slug
# contra URLs hostiles.
"""Tests del parser del slug de GitHub usado por el check de PRs
abiertas contra main (#431, #528).

CodeQL ``py/incomplete-url-substring-sanitization`` marco
``check_migrations_immutable.py:159``: la rama ``elif "github.com/" in url``
de ``parse_repo_slug`` acepta URLs cuyo path contiene ``github.com/`` pero
cuyo host no es GitHub, lo que deja al caller enviar la llamada a la API
REST contra un slug arbitrario. Caso del patron: el blueprint de
propagation reporto la misma clase de bug en PR #910 y la remediacion
subio en la cadena #528.

Estrategia: tests directos de ``parse_repo_slug`` (sin mock) con un repo
temporal y ``git remote add origin <url>`` que ejercen las tres formas:
SSH legitima, HTTPS legitima, y host con ``github.com`` en la ruta.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import check_migrations_immutable as gate

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _helpers import BaseMigrationTest, git, make_repo


def _set_remote(repo: Path, url: str) -> None:
    """Establece ``remote.origin.url`` al valor literal ``url``."""
    # ``git remote add`` falla si el remote ya existe; ``remote set-url``
    # lo actualiza sin tocar otros ajustes.
    git(repo, "remote", "add", "origin", url)
    git(repo, "remote", "set-url", "origin", url)


class ParseRepoSlugHostileURLTest(BaseMigrationTest):
    """Cubre la remediacion del finding CodeQL de PR #910 (Refs #528)."""

    def test_host_with_githubcom_in_path_returns_none(self) -> None:
        """URL ``https://evil.example/github.com/owner/repo.git``: el host
        no es GitHub, ``github.com/`` solo aparece en la ruta. El parser
        debe devolver ``None`` para que el caller falle cerrado."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "https://evil.example/github.com/owner/repo.git")
        self.assertIsNone(gate.parse_repo_slug(repo))

    def test_githubcom_subdomain_returns_none(self) -> None:
        """URL ``https://github.com.evil.com/owner/repo.git``: el host es
        ``github.com.evil.com`` (un sufijo), no ``github.com``. El match
        por substring ``"github.com/" in url`` aceptaria esta URL; el
        parser con ``urlparse`` la rechaza."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "https://github.com.evil.com/owner/repo.git")
        self.assertIsNone(gate.parse_repo_slug(repo))

    def test_legit_https_returns_slug(self) -> None:
        """URL ``https://github.com/owner/repo.git``: caso legitimo. El
        parser sigue devolviendo ``owner/repo`` tras la remediacion."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "https://github.com/owner/repo.git")
        self.assertEqual(gate.parse_repo_slug(repo), "owner/repo")

    def test_legit_https_without_dotgit_returns_slug(self) -> None:
        """URL ``https://github.com/owner/repo`` (sin ``.git``): caso
        legitimo comun. El parser normaliza el sufijo opcional."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "https://github.com/owner/repo")
        self.assertEqual(gate.parse_repo_slug(repo), "owner/repo")

    def test_legit_ssh_returns_slug(self) -> None:
        """URL ``git@github.com:owner/repo.git``: caso SSH legitimo. La
        comprobacion ``startswith("git@github.com:")`` ya era exacta
        antes de la remediacion; el parser sigue devolviendo
        ``owner/repo`` y no la rompemos."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "git@github.com:owner/repo.git")
        self.assertEqual(gate.parse_repo_slug(repo), "owner/repo")

    def test_non_github_host_returns_none(self) -> None:
        """URL ``https://gitlab.com/owner/repo.git``: el host no es
        GitHub, asi que el parser devuelve ``None`` (el caller falla
        cerrado con un mensaje claro)."""
        repo = make_repo(self.tmp, {"README.md": "hi"})
        _set_remote(repo, "https://gitlab.com/owner/repo.git")
        self.assertIsNone(gate.parse_repo_slug(repo))


if __name__ == "__main__":
    import unittest
    unittest.main()
