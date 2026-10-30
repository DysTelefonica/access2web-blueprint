"""Helper compartido de detección de Dependabot (DysTelefonica/team-skills#1113)."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from dependabot import DEPENDABOT_ACTOR, is_dependabot_branch, is_dependabot_pr  # noqa: E402


class DependabotDetectionTests(unittest.TestCase):
    def test_dependabot_actor_on_dependabot_branches_is_dependabot(self):
        for ref in ("dependabot/npm_and_yarn/lodash-4.17.21",
                    "dependabot/github_actions/actions/checkout-4",
                    "dependabot/pip/requests-2.32.0",
                    "dependabot/uv/ruff-0.6.0"):
            with self.subTest(ref=ref):
                self.assertTrue(is_dependabot_pr(DEPENDABOT_ACTOR, ref))

    def test_actor_alone_or_branch_alone_is_not_enough(self):
        self.assertFalse(is_dependabot_pr(DEPENDABOT_ACTOR, "feat/lodash"))
        self.assertFalse(is_dependabot_pr("ardelperal", "dependabot/npm_and_yarn/lodash-4.17.21"))

    def test_empty_inputs_are_not_dependabot(self):
        self.assertFalse(is_dependabot_pr("", ""))
        self.assertFalse(is_dependabot_pr(DEPENDABOT_ACTOR, ""))
        self.assertFalse(is_dependabot_branch(None))

    def test_branch_ecosystems_are_closed_list(self):
        self.assertFalse(is_dependabot_branch("dependabot/docker/alpine-3.20"))


if __name__ == "__main__":
    unittest.main()
