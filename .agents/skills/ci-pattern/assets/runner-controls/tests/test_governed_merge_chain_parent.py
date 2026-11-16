#!/usr/bin/env python3
# ci-pattern asset tests — governed merge vs an unmerged chain parent (DysTelefonica/team-skills#1597)
"""Un hijo de cadena que lleva commits de un padre abierto sin fusionar no se fusiona.

APAP_WEB#4019 (rechazado) entro en la base a traves de su hijo #4021. Solo una
PR abierta mas vieja (numero menor) puede ser el padre; los descendientes
declarados no cuentan. Si el listado o un compare fallan, falla cerrado."""
import json
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from test_governed_merge import CHECK, base_state, contract, run_gate
from test_push_compliance import REPO, SHA

PARENT, OTHER, OLD, MAIN_TIP = "b" * 40, "c" * 40, "e" * 40, "d" * 40
ACTION = ("merge or close the parent first, or rebuild the child branch from the base "
          "without the parent's commits")
LISTING = f"GET repos/{REPO}/pulls"


def open_pull(number: int, sha: str, base: str = "main", **extra: object) -> dict:
    return {"number": number, "state": "open", "head": {"sha": sha}, "base": {"ref": base}, **extra}


def child_state(*commits: str, open_pulls: list | None = None) -> dict:
    """La PR 7 (rama ``feat-7``) con ``commits`` fuera de ``main``, del mas viejo al mas nuevo.

    Una otra PR cuya cabeza es un commit del hijo se separa ahi; si no, en ``main``.
    """
    state = base_state()
    state["pr_details"]["7"]["head"].update(ref="feat-7", repo={"full_name": REPO})
    state["open_pulls"] = [open_pull(7, SHA)] if open_pulls is None else open_pulls
    state["compares"] = {f"main...{SHA}": {
        "total_commits": len(commits), "commits": [{"sha": sha} for sha in commits],
        "files": [], "behind_by": 0}}
    for item in state["open_pulls"]:
        other = item.get("head", {}).get("sha")
        if other and other != SHA:
            inside = other in commits
            state["compares"][f"{other}...{SHA}"] = {
                "merge_base_commit": {"sha": other if inside else MAIN_TIP}, "files": [], "behind_by": 0,
                "ahead_by": len(commits) - commits.index(other) - 1 if inside else len(commits)}
    return state


def fork_at(state: dict, other: str, fork: str, ahead: int) -> dict:
    state["compares"][f"{other}...{SHA}"].update(merge_base_commit={"sha": fork}, ahead_by=ahead)
    return state


class UnmergedParentTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def gate(self, state: dict):
        proc = run_gate(self.root, contract(), state)
        self.stored = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        return proc

    def assert_outcome(self, state: dict, *parents: tuple[int, str]):
        """Sin ``parents`` se fusiona; con ellos, 1 sin merge y un BLOCKED por padre."""
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1 if parents else 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.stored["merges"]), 0 if parents else 1)
        for number, sha in parents:
            self.assertIn(f"  BLOCKED  PR #7 carries the commits of open PR #{number} (head {sha[:12]}), "
                          f"not yet merged into 'main': {ACTION}", proc.stdout)
        self.assertEqual(proc.stdout.count("carries the commits"), len(parents))
        self.assertIn("VERDICT: PR #7 not merged" if parents else "MERGED: PR #7 merged", proc.stdout)
        return proc

    def assert_fails_closed(self, state: dict, message: str):
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.stored["merges"], [])
        self.assertIn(message, proc.stderr)
        return proc

    def test_s2_child_carrying_an_open_parent_head_is_not_merged(self) -> None:
        self.assert_outcome(child_state(PARENT, SHA, open_pulls=[open_pull(6, PARENT), open_pull(7, SHA)]),
                            (6, PARENT))

    def test_s2_parent_on_a_later_page_of_the_listing_is_found(self) -> None:
        state = child_state(PARENT, SHA, open_pulls=[open_pull(6, PARENT), open_pull(7, SHA)])
        state["open_pulls_pages"] = [[open_pull(7, SHA)], [open_pull(6, PARENT)]]
        self.assert_outcome(state, (6, PARENT))

    def test_s3_child_merges_once_the_parent_is_in_the_base(self) -> None:
        for pulls in ([open_pull(7, SHA)], [open_pull(6, PARENT), open_pull(7, SHA)]):
            with self.subTest(len(pulls)):
                self.assert_outcome(child_state(SHA, open_pulls=pulls))

    def test_s4_unrelated_open_prs_leave_the_merge_unchanged(self) -> None:
        cases = {"alone": None, "older PR on main": [open_pull(6, OTHER), open_pull(7, SHA)],
                 "older PR on another base": [open_pull(5, OTHER, "develop"), open_pull(7, SHA)],
                 "older PR with the child's own head": [open_pull(6, SHA, "develop"), open_pull(7, SHA)]}
        for name, pulls in cases.items():
            with self.subTest(name):
                self.assert_outcome(child_state(PARENT, SHA, open_pulls=pulls))

    def test_s5_listing_that_keeps_failing_fails_closed_without_merging(self) -> None:
        state = child_state(PARENT, SHA)
        state["open_pulls_transient"] = 99
        self.assert_fails_closed(state, "(falló tras 4 intentos)")
        self.assertEqual(self.stored["calls"].count(LISTING), 4)

    def test_s5_a_transient_listing_failure_is_retried(self) -> None:
        state = child_state(PARENT, SHA, open_pulls=[open_pull(6, PARENT), open_pull(7, SHA)])
        state["open_pulls_transient"] = 1
        self.assert_outcome(state, (6, PARENT))
        self.assertEqual(self.stored["calls"].count(LISTING), 2)

    def test_s5_a_403_on_the_listing_fails_closed_after_one_call(self) -> None:
        state = child_state(PARENT, SHA)
        state["open_pulls_forbidden"] = True
        self.assert_fails_closed(state, "fail-closed: GitHub API")
        self.assertEqual(self.stored["calls"].count(LISTING), 1)

    def test_s1_parents_on_another_base_draft_or_two_block(self) -> None:
        cases = {"other base": [open_pull(6, PARENT, "feat-a")], "draft": [open_pull(6, PARENT, draft=True)],
                 "two parents": [open_pull(5, OLD), open_pull(6, PARENT)]}
        for name, pulls in cases.items():
            with self.subTest(name):
                parents = [(item["number"], item["head"]["sha"]) for item in pulls]
                self.assert_outcome(child_state(OLD, PARENT, SHA, open_pulls=[*pulls, open_pull(7, SHA)]),
                                    *parents)

    def test_a_reversed_chain_blocks_the_newer_child_only(self) -> None:
        """B (b1, b2) y C cortado de b1: C (PR mas nueva) nunca bloquea a B; B si bloquea a C."""
        for base in ("main", "feat-7"):
            with self.subTest(base):
                self.assert_outcome(fork_at(child_state(OLD, SHA, open_pulls=[
                    open_pull(7, SHA), open_pull(8, OTHER, base)]), OTHER, OLD, 1))
                self.assertNotIn(f"GET repos/{REPO}/compare/{OTHER}...{SHA}", self.stored["calls"])
        self.assert_outcome(fork_at(child_state(OLD, SHA, open_pulls=[
            open_pull(6, PARENT), open_pull(7, SHA)]), PARENT, OLD, 1), (6, PARENT))

    def test_b_declared_descendant_and_contained_pr_do_not_block(self) -> None:
        for base, ahead in (("feat-7", 1), ("feat-b", 0)):
            with self.subTest(base):
                self.assert_outcome(fork_at(child_state(OLD, SHA, open_pulls=[
                    open_pull(5, OTHER, base), open_pull(7, SHA)]), OTHER, OLD if ahead else SHA, ahead))

    def test_c_parent_head_among_the_child_commits_blocks_without_compare(self) -> None:
        self.assert_outcome(child_state(PARENT, SHA, open_pulls=[open_pull(6, PARENT)]), (6, PARENT))
        self.assertNotIn(f"GET repos/{REPO}/compare/{PARENT}...{SHA}", self.stored["calls"])

    def test_e_another_finding_skips_the_parent_check(self) -> None:
        state = child_state(PARENT, SHA, open_pulls=[open_pull(6, PARENT)])
        state["check_runs"][SHA][0]["conclusion"] = "failure"
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(CHECK, proc.stdout)
        self.assertNotIn(LISTING, self.stored["calls"])

    def test_f_failed_compare_fails_closed_naming_the_other_pr(self) -> None:
        state = child_state(PARENT, SHA, open_pulls=[open_pull(6, OTHER)])
        state["compare_errors"] = {f"{OTHER}...{SHA}": 404}
        self.assert_fails_closed(state, "the compare with open PR #6 failed (fail closed)")

    def test_s5_open_pull_without_head_fails_closed(self) -> None:
        self.assert_fails_closed(child_state(SHA, open_pulls=[{"number": 6, "base": {"ref": "main"}}]),
                                 "the open pulls listing is not the API pulls shape")

    def test_s5_ill_shaped_compare_with_another_pull_fails_closed(self) -> None:
        cases = {"no merge base": {"merge_base_commit": None}, "bool ahead_by": {"ahead_by": True},
                 "short merge base sha": {"merge_base_commit": {"sha": "abc"}}}
        for name, change in cases.items():
            with self.subTest(name):
                state = child_state(SHA, open_pulls=[open_pull(6, OTHER)])
                state["compares"][f"{OTHER}...{SHA}"].update(change)
                self.assert_fails_closed(state, f"open PR #6: compare {OTHER}...{SHA} has no valid "
                                                "'ahead_by' and 'merge_base_commit'")

    def test_incomplete_commit_list_of_the_child_fails_closed(self) -> None:
        state = child_state(SHA, open_pulls=[open_pull(6, PARENT)])
        state["compares"][f"main...{SHA}"]["total_commits"] = 300
        self.assert_fails_closed(state, f"compare main...{SHA} lists 1 of 300 commits")


class ComparePagingTests(unittest.TestCase):
    """El compare del hijo se pagina (un tope sin paginar falla cerrado); el de otra PR pide una pagina de 1."""

    @staticmethod
    def gate():
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
        import check_governed_merge
        return check_governed_merge

    def ahead(self, total: int, pages) -> set[str]:
        gate = self.gate()
        shas = [f"{i:040x}" for i in range(total)]

        def fake(path: str) -> dict:
            query = parse_qs(urlsplit(path).query)
            page = pages(shas, int(query["page"][0]), int(query["per_page"][0]))
            return {"total_commits": total, "commits": [{"sha": sha} for sha in page]}
        with unittest.mock.patch.object(gate, "gh_json", fake):
            return gate.commits_ahead("o/r", "main", SHA)

    def test_more_than_250_commits_are_paged(self) -> None:
        self.assertEqual(len(self.ahead(251, lambda shas, page, size: shas[(page - 1) * size:page * size])), 251)

    def test_api_cap_without_paging_fails_closed(self) -> None:
        with self.assertRaises(self.gate().FailClosedError):
            self.ahead(300, lambda shas, page, size: shas[:250])

    def test_d_other_pr_compare_reads_one_commit_page(self) -> None:
        gate = self.gate()
        fake = unittest.mock.Mock(return_value={"ahead_by": 1, "merge_base_commit": {"sha": OLD}})
        with unittest.mock.patch.object(gate, "get_json", fake):
            self.assertEqual(gate.fork_point("o/r", 6, PARENT, SHA), (1, OLD))
        fake.assert_called_once_with(f"repos/o/r/compare/{PARENT}...{SHA}?per_page=1")


if __name__ == "__main__":
    unittest.main()
