#!/usr/bin/env python3
"""#1190: la punta de una cadena nueva lista en `Closes` cada issue que sus eslabones referencian.

Aviso, no bloqueo (HR-7, #1172). Solo cadenas nuevas: la punta declara `- chain: <id>`
en su Chain Context. stdlib exclusivamente.
"""

import importlib.machinery
import importlib.util
import os
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_contract.py"

_spec = importlib.util.spec_from_loader(
    "pr_contract_gate_chain", importlib.machinery.SourceFileLoader("pr_contract_gate_chain", str(GATE)))
pr_contract = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pr_contract)

REPO = "DysTelefonica/example"
TIP_BODY = (
    "## Chain Context\n\n"
    "- chain: cadena-1190\n"
    "- position: 3/3\n"
    "- base: main\n\n"
    "Closes #10\n"
)


def _sibling(number, refs, chain="cadena-1190"):
    """PR eslabón: su Chain Context lleva el id de la cadena y sus Refs."""
    body = f"## Chain Context\n\n- chain: {chain}\n- position: 1/3\n\n"
    body += "".join(f"Refs #{n}\n" for n in refs)
    return {"number": number, "body": body}


class ChainClosesFindingsTests(unittest.TestCase):
    def test_tip_missing_a_link_issue_gets_one_advisory_naming_it(self):
        siblings = [_sibling(1, [10]), _sibling(2, [11])]

        findings = pr_contract.chain_closes_findings(TIP_BODY, siblings, {})

        self.assertEqual(len(findings), 1)
        self.assertIn("#11", findings[0])
        self.assertNotIn("#10", findings[0])
        self.assertIn("cadena-1190", findings[0])

    def test_tip_closing_every_link_issue_gets_no_advisory(self):
        body = TIP_BODY + "Closes #11\n"
        siblings = [_sibling(1, [10]), _sibling(2, [11])]

        self.assertEqual(pr_contract.chain_closes_findings(body, siblings, {}), [])

    def test_a_tip_without_chain_id_is_not_checked(self):
        """Solo cadenas nuevas: sin `- chain:` la punta no se mide."""
        body = "## Chain Context\n\n- position: 3/3\n\nClose #10\n"
        siblings = [_sibling(1, [11])]

        self.assertEqual(pr_contract.chain_closes_findings(body, siblings, {}), [])

    def test_links_of_another_chain_are_not_counted(self):
        siblings = [_sibling(1, [11], chain="otra-cadena")]

        self.assertEqual(pr_contract.chain_closes_findings(TIP_BODY, siblings, {}), [])

    def test_a_closing_word_in_a_link_does_not_count_as_a_reference(self):
        """Un eslabón con `Closes` ya tiene su propio aviso (HR-7); aquí solo cuentan los `Refs`."""
        sibling = {"number": 1, "body": "- chain: cadena-1190\n\nCloses #11\n"}

        self.assertEqual(pr_contract.chain_closes_findings(TIP_BODY, [sibling], {}), [])


class ChainClosesCheckTests(unittest.TestCase):
    def _event(self):
        return {"repository": {"full_name": REPO}}

    def _pr(self):
        return {"number": 3, "body": TIP_BODY, "labels": []}

    def test_api_failure_is_an_advisory_not_a_finding(self):
        advisories = []

        def failing(*_args):
            raise pr_contract._GITHUB_REST.GitHubApiError("HTTP 503")

        old = os.environ.get("GH_TOKEN")
        os.environ["GH_TOKEN"] = "x"
        try:
            pr_contract.check_chain_closes(self._pr(), self._event(), {}, advisories, fetch=failing)
        finally:
            if old is None:
                os.environ.pop("GH_TOKEN", None)
            else:
                os.environ["GH_TOKEN"] = old

        self.assertEqual(len(advisories), 1)
        self.assertIn("no se pudo leer la cadena", advisories[0])

    def test_without_token_the_check_is_skipped_silently(self):
        advisories = []
        saved = {k: os.environ.pop(k) for k in ("GH_TOKEN", "GITHUB_TOKEN") if k in os.environ}
        try:
            pr_contract.check_chain_closes(
                self._pr(), self._event(), {}, advisories,
                fetch=lambda *_args: self.fail("no debe leer la API sin token"))
        finally:
            os.environ.update(saved)

        self.assertEqual(advisories, [])

    def test_siblings_come_from_the_fetch_for_this_repo_and_tip(self):
        advisories = []
        seen = {}

        def fetch(repo, number):
            seen["repo"], seen["number"] = repo, number
            return [_sibling(2, [11])]

        old = os.environ.get("GH_TOKEN")
        os.environ["GH_TOKEN"] = "x"
        try:
            pr_contract.check_chain_closes(self._pr(), self._event(), {}, advisories, fetch=fetch)
        finally:
            if old is None:
                os.environ.pop("GH_TOKEN", None)
            else:
                os.environ["GH_TOKEN"] = old

        self.assertEqual(seen, {"repo": REPO, "number": 3})
        self.assertEqual(len(advisories), 1)
        self.assertIn("#11", advisories[0])

    def test_fetch_leaves_the_tip_out_of_the_listing(self):
        listing = [{"number": 3, "body": TIP_BODY}, _sibling(2, [11])]
        original = pr_contract._GITHUB_REST.get_json
        pr_contract._GITHUB_REST.get_json = lambda _path: listing
        try:
            siblings = pr_contract.fetch_chain_siblings(REPO, 3)
        finally:
            pr_contract._GITHUB_REST.get_json = original

        self.assertEqual([s["number"] for s in siblings], [2])


if __name__ == "__main__":
    unittest.main()
