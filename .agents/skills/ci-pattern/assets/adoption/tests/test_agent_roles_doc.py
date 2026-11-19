"""#1768/#1772: the operating doc and the agents block without agent roles.

A consumer without `agent_roles: enabled` in its `.team-skills.yaml` gets a
role-free operating doc: the governed merge is launched by the agent with the
profile of the administrator running the session, with CI green, and nobody
needs a separate registered order. With `agent_roles: enabled` the doc keeps
today's text. The same rule applies to the `ci-pattern-agents` block, which
the consumer gate compares with `_agents_block_for`.
"""
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
_spec = importlib.util.spec_from_file_location("adoption_check_helpers", HERE / "test_adoption_check.py")
helpers = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(helpers)

ROLE_WORDS = ("decide un humano", "mantenedor", "coordinador", "revisor")
ADMIN = "administrador que lleva la sesión"


def generate(repo: Path) -> str:
    out = repo.parent / "doc.md"
    code, _, err = helpers.run_cli(["adoption", "generate-doc", str(repo), "--out", str(out)])
    if code != 0:
        raise AssertionError(err)
    return out.read_text(encoding="utf-8")


class OperatingDocWithoutRolesTests(unittest.TestCase):
    def test_default_doc_names_no_role_actor(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            text = generate(helpers.make_docs_repo(tmp))
        for word in ROLE_WORDS:
            self.assertNotIn(word, text, f"role-free doc still says {word!r}")
        self.assertIn(f"perfil del {ADMIN}", text)
        self.assertIn("CI en verde", text)

    def test_roles_enabled_keeps_todays_text(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = helpers.make_docs_repo(tmp)
            (repo / ".team-skills.yaml").write_text("project_type: web\nagent_roles: enabled\n", encoding="utf-8")
            text = generate(repo)
        self.assertIn("El merge lo decide un humano con CI en verde", text)
        self.assertIn("pregunte en la issue o al mantenedor", text)

    def test_agents_block_follows_the_declaration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = helpers.make_docs_repo(tmp)
            off = helpers.CLI._agents_block_for(repo)
            (repo / ".team-skills.yaml").write_text("agent_roles: enabled\n", encoding="utf-8")
            on = helpers.CLI._agents_block_for(repo)
        self.assertNotIn("mantenedor", off)
        self.assertIn(f"pregunte en la issue o al {ADMIN}", off)
        self.assertEqual(on, helpers.CLI.OPERATING_DOC_AGENTS_BLOCK)


if __name__ == "__main__":
    unittest.main()
