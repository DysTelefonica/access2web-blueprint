#!/usr/bin/env python3
"""HR-93 (#1726): el tipo de la PR sale de su título convencional.

Un test por criterio de la issue #1726: C1 a C4.
"""

import importlib.util
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_contract.py"
spec = importlib.util.spec_from_file_location("check_pr_contract_title", GATE)
MOD = importlib.util.module_from_spec(spec)
spec.loader.exec_module(MOD)

PLAYBOOK = "docs/codebase/maintainer-playbook.md"


class PrTitleTests(unittest.TestCase):
    """El tipo de la PR se deriva del título; la etiqueta `type:*` es opcional."""

    def _judge(self, pr):
        findings, advisories = [], []
        ptype = MOD.check_pr_title(pr, findings, advisories)
        return ptype, findings, advisories

    def _playbook(self, tmp, sections):
        repo = Path(tmp)
        playbook = repo / PLAYBOOK
        playbook.parent.mkdir(parents=True)
        playbook.write_text("".join(f"## {name}\n\n" for name in sections),
                            encoding="utf-8")
        return repo

    def test_conventional_title_without_type_label_passes(self):
        """C1: título convencional válido, sin etiqueta `type:*`: nada bloquea."""
        pr = {"title": "feat(pr-contract): el tipo sale del título", "labels": []}
        ptype, findings, advisories = self._judge(pr)
        self.assertEqual(ptype, "feat")
        self.assertEqual((findings, advisories), ([], []))

    def test_non_conventional_title_blocks_naming_the_rule(self):
        """C2: título no convencional: bloquea y nombra la regla y los tipos."""
        pr = {"title": "el tipo sale del título", "labels": []}
        ptype, findings, advisories = self._judge(pr)
        self.assertIsNone(ptype)
        self.assertEqual(len(findings), 1)
        self.assertIn("HR-93", findings[0])
        self.assertIn("feat", findings[0])
        self.assertIn("el tipo sale del título", findings[0])
        self.assertEqual(advisories, [])

    def test_conflicting_type_label_warns_and_the_title_wins(self):
        """C3: etiqueta `type:*` que contradice el título: avisa y manda el título."""
        pr = {"title": "fix(pr-contract): el tipo sale del título",
              "labels": [{"name": "type:feature"}]}
        ptype, findings, advisories = self._judge(pr)
        self.assertEqual(ptype, "fix")
        self.assertEqual(findings, [])
        self.assertEqual(len(advisories), 1)
        self.assertIn("type:feature", advisories[0])
        self.assertIn("manda el título", advisories[0])

    def test_the_equivalent_label_does_not_warn(self):
        """C3: `type:feature` con `feat` y `type:bug` con `fix` son el mismo tipo."""
        for title, label in (("feat: x", "type:feature"), ("fix: x", "type:bug")):
            ptype, findings, advisories = self._judge(
                {"title": title, "labels": [{"name": label}]})
            self.assertEqual(ptype, title.split(":")[0])
            self.assertEqual((findings, advisories), ([], []), title)

    def test_checklist_is_selected_by_the_title_type(self):
        """C4: HR-64 exige la checklist del tipo del título, no la de la etiqueta.

        El tipo del título manda con cualquiera de los dos vocabularios: el
        canónico (`type:docs`) y el de los hosts (`type:feature`, `type:bug`)."""
        cases = (("docs(pr-contract): el tipo sale del título", "type:feature",
                  {"type:feature": "#feature", "type:docs": "#docs"}, "#docs"),
                 ("fix(pr-contract): el tipo sale del título", "type:feature",
                  {"type:feature": "#feature", "type:bug": "#bug"}, "#bug"))
        for title, label, checklists, expected in cases:
            with self.subTest(title=title), tempfile.TemporaryDirectory() as tmp:
                repo = self._playbook(tmp, ["feature", "docs", "bug"])
                policy = {"docs_evidence": {
                    "pages": [{"page": PLAYBOOK, "covers": ["app/*"]}],
                    "checklists": {key: PLAYBOOK + anchor
                                   for key, anchor in checklists.items()}}}
                pr = {"title": title, "labels": [{"name": label}], "body": ""}
                advisories = []
                MOD.check_docs_evidence(pr, str(repo), [], policy, advisories)
                joined = "\n".join(advisories)
                self.assertIn(expected, joined)
                self.assertNotIn(checklists["type:feature"], joined)

    def test_a_dependabot_pr_is_exempt(self):
        """El título de un PR de Dependabot lo genera GitHub: no se mide."""
        pr = {"title": "Bump ruff from 0.15.20 to 0.15.21", "labels": [],
              "user": {"login": "dependabot[bot]"},
              "head": {"ref": "dependabot/pip/ruff-0.15.21"}}
        ptype, findings, advisories = self._judge(pr)
        self.assertIsNone(ptype)
        self.assertEqual((findings, advisories), ([], []))

    def test_an_empty_title_blocks_and_a_missing_one_is_not_measured(self):
        """Un PR sin título es un defecto; un payload sin título no se mide."""
        ptype, findings, advisories = self._judge({"title": "", "labels": []})
        self.assertIsNone(ptype)
        self.assertEqual(len(findings), 1)
        self.assertIn("HR-93", findings[0])
        self.assertEqual(advisories, [])
        ptype, findings, advisories = self._judge({"labels": []})
        self.assertIsNone(ptype)
        self.assertEqual((findings, advisories), ([], []))


if __name__ == "__main__":
    unittest.main()
