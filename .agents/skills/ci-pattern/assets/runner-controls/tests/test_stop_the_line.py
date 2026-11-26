#!/usr/bin/env python3
"""Tests unitarios de stop_the_line.py (#1627): la regla, sin red.

El lector es falso; el gate recibe una función que devuelve los enlaces del PR.
"""

import base64
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

ASSET_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ASSET_DIR))

import stop_the_line as stl  # noqa: E402
from check_push_compliance import FailClosedError  # noqa: E402
from github_rest import GitHubApiError  # noqa: E402

REPO = "owner/repo"
WORKFLOW = "tests.yml"
RUN_URL = "https://github.com/owner/repo/actions/runs/777"
NOT_FOUND = GitHubApiError("HTTP 404", 404)
SERVER_ERROR = GitHubApiError("HTTP 500", 500)
REDISPATCH = "gh workflow run tests.yml --repo owner/repo --ref main"
MAIN_RED = "prerelease:main-red"


def profile_doc(content):
    raw = content if isinstance(content, str) else json.dumps(content)
    return {"encoding": "base64", "content": base64.b64encode(raw.encode("utf-8")).decode("ascii")}


def night(conclusion="failure", event="schedule", status="completed", created_at="",
          display_title=""):
    return {"id": 777, "event": event, "status": status, "conclusion": conclusion,
            "html_url": RUN_URL, "created_at": created_at, "display_title": display_title}


class FakeReader:
    """Contrato del lector; registra las ramas con las que se le llama."""

    def __init__(self, default="main", profile=None, profile_error=None, runs=(),
                 runs_error=None, issues=(), issues_error=None, default_error=None,
                 main_red=(), main_red_error=None):
        self.default, self.default_error = default, default_error
        self.profile, self.profile_error = profile, profile_error
        self.runs, self.runs_error = list(runs), runs_error
        self.issues, self.issues_error = list(issues), issues_error
        self.main_red, self.main_red_error = list(main_red), main_red_error
        self.branches = []

    def default_branch(self, repo):
        if self.default_error is not None:
            raise self.default_error
        return self.default

    def contents(self, repo, path, ref):
        self.branches.append(ref)
        if self.profile_error is not None:
            raise self.profile_error
        if self.profile is None:
            raise NOT_FOUND
        return self.profile

    def workflow_runs(self, repo, workflow, branch, event):
        self.branches.append(branch)
        if self.runs_error is not None:
            raise self.runs_error
        return {"workflow_runs": [run for run in self.runs if run.get("event") == event]}

    def open_issues(self, repo, label):
        if label == MAIN_RED:
            if self.main_red_error is not None:
                raise self.main_red_error
            return self.main_red
        if self.issues_error is not None:
            raise self.issues_error
        return self.issues


def issue(number):
    return {"number": number, "state": "open"}


def evaluate(reader, linked=(), target=None):
    target = target or stl.Target(repo=REPO, pr=42, nightly_workflow=WORKFLOW)
    return stl.evaluate(target, reader, lambda: list(linked))


class TestTheRule(unittest.TestCase):
    def test_red_night_with_no_labelled_issue_passes(self):
        self.assertEqual(evaluate(FakeReader(runs=[night()])), [])

    def test_red_night_blocks_a_pr_that_links_no_red_issue_without_a_profile(self):
        reader = FakeReader(runs=[night()], issues=[issue(11), issue(12)])
        findings = evaluate(reader, linked=[(REPO, 99)])
        self.assertEqual(len(findings), 1)
        self.assertIn(RUN_URL, findings[0])
        self.assertIn("#11, #12", findings[0])

    def test_block_message_gives_the_redispatch_on_the_default_branch(self):
        reader = FakeReader(runs=[night("cancelled")], issues=[issue(11)])
        self.assertIn(REDISPATCH, evaluate(reader)[0])

    def test_a_pr_that_links_a_red_issue_of_the_same_repo_passes(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)])
        self.assertEqual(evaluate(reader, linked=[("Owner/Repo", 11)]), [])

    def test_cross_repo_link_with_the_same_number_is_not_a_fix(self):
        reader = FakeReader(runs=[night()], issues=[issue(12)])
        self.assertEqual(len(evaluate(reader, linked=[("owner/other", 12)])), 1)

    def test_green_night_passes(self):
        reader = FakeReader(runs=[night("success")], issues=[issue(11)])
        self.assertEqual(evaluate(reader), [])

    def test_a_light_green_dispatch_does_not_lift_a_red_schedule(self):
        """#1702: un dispatch ligero en verde (sin --dispatch-marker) no cuenta como noche."""
        reader = FakeReader(runs=[
            night("success", "workflow_dispatch", created_at="2026-11-14T02:00:00Z", display_title="suite=light"),
            night("failure", created_at="2026-11-14T01:00:00Z")], issues=[issue(11)])
        self.assertEqual(len(evaluate(reader)), 1)

    def test_a_green_dispatch_with_the_marker_in_its_title_lifts_the_line(self):
        """#1702: con --dispatch-marker suite=full, un dispatch en verde con ese título sí cuenta."""
        reader = FakeReader(runs=[
            night("success", "workflow_dispatch", created_at="2026-11-14T02:00:00Z", display_title="suite=full"),
            night("failure", created_at="2026-11-14T01:00:00Z")], issues=[issue(11)])
        target = stl.Target(repo=REPO, pr=42, nightly_workflow=WORKFLOW, dispatch_marker="suite=full")
        self.assertEqual(evaluate(reader, target=target), [])

    def test_a_green_dispatch_without_the_marker_in_its_title_does_not_lift_the_line(self):
        """#1702: con --dispatch-marker suite=full, un dispatch en verde sin ese título no cuenta."""
        reader = FakeReader(runs=[
            night("success", "workflow_dispatch", created_at="2026-11-14T02:00:00Z", display_title="suite=light"),
            night("failure", created_at="2026-11-14T01:00:00Z")], issues=[issue(11)])
        target = stl.Target(repo=REPO, pr=42, nightly_workflow=WORKFLOW, dispatch_marker="suite=full")
        self.assertEqual(len(evaluate(reader, target=target)), 1)

    def test_an_in_progress_run_is_skipped_for_the_last_completed_one(self):
        red_then_running = FakeReader(issues=[issue(11)], runs=[
            night(None, status="in_progress"), night("failure")])
        self.assertEqual(len(evaluate(red_then_running)), 1)
        green_dispatch_after_red = FakeReader(issues=[issue(11)], runs=[
            night(None, status="in_progress"), night("failure"),
            night("success", event="workflow_dispatch", display_title="suite=light")])
        self.assertEqual(len(evaluate(green_dispatch_after_red)), 1)

    def test_pull_requests_are_not_counted_as_red_issues(self):
        pr_entry = {"number": 50, "pull_request": {"url": "u"}}
        reader = FakeReader(runs=[night()], issues=[pr_entry, issue(11)])
        self.assertIn("#11", evaluate(reader)[0])
        self.assertNotIn("#50", evaluate(reader)[0])

    def test_every_read_uses_the_default_branch_not_the_pr_base(self):
        reader = FakeReader(default="trunk", profile=profile_doc({"enabled": True}),
                            runs=[night()], issues=[issue(11)])
        evaluate(reader)
        self.assertEqual(set(reader.branches), {"trunk"})

    def test_linked_refs_are_read_only_when_there_is_red(self):
        calls = []
        stl.evaluate(stl.Target(REPO, 42, WORKFLOW), FakeReader(runs=[night("success")]),
                     lambda: calls.append(1) or [])
        self.assertEqual(calls, [])

    def test_a_failed_linked_read_propagates(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)])

        def broken():
            raise FailClosedError("links unreadable")

        with self.assertRaises(FailClosedError):
            stl.evaluate(stl.Target(REPO, 42, WORKFLOW), reader, broken)

    def test_label_is_url_encoded_in_the_listing_path(self):
        with mock.patch.object(stl, "get_json", return_value=[]) as get:
            stl.GitHubReader().open_issues(REPO, "prerelease:night red")
        self.assertIn("labels=prerelease%3Anight%20red", get.call_args.args[0])


class TestProfile(unittest.TestCase):
    def test_missing_nightly_workflow_with_a_red_label_holds_naming_the_option(self):
        """#1702: un workflow nocturno inexistente sale con 2 y nombra el workflow y --nightly-workflow."""
        reader = FakeReader(issues=[issue(11)], runs_error=NOT_FOUND)
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(reader)
        self.assertIn(WORKFLOW, str(ctx.exception))
        self.assertIn("--nightly-workflow", str(ctx.exception))

    def test_profile_on_without_a_completed_run_holds(self):
        reader = FakeReader(profile=profile_doc({"enabled": True}), issues=[issue(11)], runs=[])
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(reader)
        self.assertIn(REDISPATCH, str(ctx.exception))

    def test_profile_enabled_false_still_holds_when_the_nightly_workflow_is_missing(self):
        """#1702: el 404 del workflow nocturno retiene aunque el perfil esté apagado."""
        reader = FakeReader(profile=profile_doc({"enabled": False}), issues=[issue(11)],
                            runs_error=NOT_FOUND)
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(reader)
        self.assertIn("--nightly-workflow", str(ctx.exception))

    def test_unreadable_profile_holds_with_the_redispatch_command(self):
        reader = FakeReader(profile_error=SERVER_ERROR)
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(reader)
        self.assertIn(REDISPATCH, str(ctx.exception))

    def test_invalid_profile_holds(self):
        reader = FakeReader(profile=profile_doc({"enabled": "yes"}))
        with self.assertRaises(FailClosedError):
            evaluate(reader)


class TestFailClosed(unittest.TestCase):
    def assert_holds(self, reader):
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(reader)
        self.assertIn(REDISPATCH, str(ctx.exception))

    def test_default_branch_read_error_holds(self):
        self.assert_holds(FakeReader(default_error=SERVER_ERROR))

    def test_labelled_issue_read_error_holds(self):
        self.assert_holds(FakeReader(issues_error=SERVER_ERROR))

    def test_run_listing_error_holds_when_there_is_a_signal(self):
        self.assert_holds(FakeReader(issues=[issue(11)], runs_error=SERVER_ERROR))

    def test_non_list_issues_listing_holds(self):
        reader = FakeReader()
        reader.open_issues = lambda repo, label: {"not": "a list"}
        self.assert_holds(reader)


MAIN_RED_HINT = "arregle main en el PR que enlaza #21; el siguiente run verde de main la cierra"


class TestMainRedSignal(unittest.TestCase):
    """#1708: una issue abierta `prerelease:main-red` bloquea todo PR que no la enlace.

    Es una segunda señal, independiente de la nocturna: ambas se evalúan."""

    def test_open_main_red_issue_blocks_a_pr_that_links_none(self):
        findings = evaluate(FakeReader(main_red=[issue(21)]), linked=[(REPO, 99)])
        self.assertEqual(len(findings), 1)
        self.assertIn("#21", findings[0])
        self.assertIn(MAIN_RED, findings[0])
        self.assertIn(MAIN_RED_HINT, findings[0])

    def test_the_linked_fix_passes(self):
        self.assertEqual(evaluate(FakeReader(main_red=[issue(21)]), linked=[(REPO, 21)]), [])

    def test_no_open_main_red_issue_passes(self):
        self.assertEqual(evaluate(FakeReader(main_red=[])), [])

    def test_a_same_number_in_another_repository_is_not_a_fix(self):
        findings = evaluate(FakeReader(main_red=[issue(21)]), linked=[("other/repo", 21)])
        self.assertEqual(len(findings), 1)

    def test_linked_refs_are_read_only_when_main_is_red(self):
        calls = []
        stl.evaluate(stl.Target(repo=REPO, pr=42, nightly_workflow=WORKFLOW), FakeReader(),
                     lambda: calls.append("read") or [])
        self.assertEqual(calls, [])

    def test_main_red_read_error_holds_naming_the_label(self):
        with self.assertRaises(FailClosedError) as ctx:
            evaluate(FakeReader(main_red_error=SERVER_ERROR))
        self.assertIn(MAIN_RED, str(ctx.exception))

    def test_night_and_main_red_together_both_block(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)], main_red=[issue(21)])
        findings = evaluate(reader, linked=[(REPO, 99)])
        self.assertEqual(len(findings), 2)
        self.assertTrue(any("#11" in f for f in findings))
        self.assertTrue(any("#21" in f for f in findings))

    def test_linking_only_the_night_issue_does_not_lift_main_red(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)], main_red=[issue(21)])
        findings = evaluate(reader, linked=[(REPO, 11)])
        self.assertEqual(len(findings), 1)
        self.assertIn("#21", findings[0])

    def test_linking_only_the_main_red_issue_does_not_lift_the_night(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)], main_red=[issue(21)])
        findings = evaluate(reader, linked=[(REPO, 21)])
        self.assertEqual(len(findings), 1)
        self.assertIn("#11", findings[0])

    def test_linking_both_passes(self):
        reader = FakeReader(runs=[night()], issues=[issue(11)], main_red=[issue(21)])
        self.assertEqual(evaluate(reader, linked=[(REPO, 11), (REPO, 21)]), [])

    def test_green_night_does_not_lift_an_open_main_red_issue(self):
        reader = FakeReader(runs=[night("success")], main_red=[issue(21)])
        self.assertEqual(len(evaluate(reader, linked=[])), 1)


if __name__ == "__main__":
    unittest.main()
