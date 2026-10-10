#!/usr/bin/env python3
# ci-pattern asset tests — ratchet gate, stable identities + baselines (DysTelefonica/team-skills#194)
"""Executable suite for the ratchet gate (black-box, like CI).

Core rules under test: a new finding identity not in the baseline is a
finding (HR-15); an identity eliminated from findings demands the baseline
be shrunk to lock it in (HR-15); every baseline entry carries target and
target_date and an expired date is a finding; identities are stable
(HR-16) — reordering findings or changing mutable fields like line numbers
never creates or hides an identity. Liveness: missing baseline or an empty
findings list exits 2 (HR-3: a gate that did not measure never says OK).
Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_ratchet.py"
TODAY = "2026-10-04"


def ident(rule="lint-rule", path="src/app.py", fingerprint="fp-1"):
    return {"rule": rule, "path": path, "fingerprint": fingerprint}


def entry(i=None, target="fix before 2027", target_date="2026-12-31", **ident_over):
    return {"identity": {**(i or ident()), **ident_over},
            "target": target, "target_date": target_date}


def baseline(entries, **extra):
    doc = {"baseline_version": 1, "entries": list(entries)}
    doc.update(extra)
    return doc


def findings(items, measured=True, subjects=None):
    return {"measured": measured,
            "subjects": subjects if subjects is not None else max(len(items), 1),
            "findings": list(items)}


def write_tmp(d, name, doc):
    p = Path(d, name)
    if isinstance(doc, str):
        p.write_text(doc, encoding="utf-8")
    else:
        p.write_text(json.dumps(doc), encoding="utf-8")
    return p


class GateCase(unittest.TestCase):
    def run_gate(self, findings_doc, baseline_doc, extra=()):
        with tempfile.TemporaryDirectory() as d:
            f = write_tmp(d, "findings.json", findings_doc)
            b = write_tmp(d, "baseline.json", baseline_doc)
            out = subprocess.run(
                [sys.executable, str(GATE), "--findings", str(f),
                 "--baseline", str(b), "--today", TODAY, *extra],
                capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def table(self, cases):
        for label, fdoc, bdoc, code, needle in cases:
            with self.subTest(case=label):
                got, out = self.run_gate(fdoc, bdoc)
                self.assertEqual(got, code, out)
                if needle:
                    self.assertIn(needle, out)


class InvalidInput(GateCase):
    """Exit 2: unreadable, malformed, or an empty measurement without liveness proof."""
    CASES = [
        ("findings without liveness proof", {"findings": []}, baseline([entry()]), None),
        ("findings measured false", findings([], measured=False), baseline([entry()]), None),
        ("findings subjects zero", findings([], subjects=0), baseline([entry()]), None),
        ("findings subjects negative", findings([], subjects=-1), baseline([entry()]), None),
        ("findings subjects not an int", findings([], subjects="3"), baseline([entry()]), None),
        ("findings not an object", "[1,2]", baseline([entry()]), None),
        ("findings unknown top-level key", {**findings([ident()]), "x": 1}, baseline([entry()]), None),
        ("findings entry not an object", {"measured": True, "subjects": 1, "findings": ["x"]}, baseline([entry()]), None),
        ("findings entry missing rule", {"findings": [{"path": "p", "fingerprint": "f"}]}, baseline([entry()]), None),
        ("findings entry empty fingerprint", {"findings": [ident(fingerprint="")]}, baseline([entry()]), None),
        ("duplicate findings identities", {"findings": [ident(), ident()]}, baseline([entry()]), None),
        ("baseline missing file", findings([ident()]), "not json at all", None),
        ("baseline without version", findings([ident()]), {"entries": []}, None),
        ("baseline unknown version", findings([ident()]), baseline([entry()], baseline_version=2), None),
        ("baseline entries not a list", findings([ident()]), baseline("x"), None),
        ("baseline entry missing target_date", findings([ident()]),
         {"baseline_version": 1, "entries": [{"identity": ident(), "target": "t"}]}, None),
        ("baseline entry unknown key", findings([ident()]), baseline([entry(note="x")]), None),
        ("baseline identity missing fingerprint", findings([ident()]),
         baseline([entry({"rule": "r", "path": "p"})]), None),
        ("baseline bad date", findings([ident()]), baseline([entry(target_date="31-12-2026")]), None),
        ("baseline duplicate identities", findings([ident()]), baseline([entry(), entry()]), None),
        ("baseline unknown top-level key", findings([ident()]), baseline([entry()], extra_key=1), None),
        ("relaxation without verifiable identity", findings([ident()]),
         baseline([entry()], relaxation={"reason": "r", "reference": "x"}), None),
        ("baseline empty string target", findings([ident()]), baseline([entry(target="")]), None),
    ]

    def test_exit_two(self):
        self.table([(l, f, b, 2, None) for l, f, b, _ in self.CASES])

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


class NewIdentityIsAFinding(GateCase):
    """HR-15: the ratchet rejects what is new; only the baseline can hold identities."""
    def test_findings(self):
        self.table([
            ("new identity not in baseline", findings([ident(), ident(fingerprint="fp-2")]),
             baseline([entry()]), 1, "HR-15"),
            ("new rule same fingerprint", findings([ident(rule="other-rule")]),
             baseline([entry()]), 1, "HR-15"),
            ("new path same fingerprint", findings([ident(path="src/other.py")]),
             baseline([entry()]), 1, "HR-15"),
            ("empty baseline everything is new", findings([ident()]),
             baseline([]), 1, "HR-15"),
        ])

    def test_negatives(self):
        self.table([
            ("clean exact match", findings([ident()]), baseline([entry()]), 0, None),
            ("identity survives reorder", findings([{"line": 99, **ident()}]),
             baseline([entry()]), 0, None),
            ("extra findings keys are tolerated", findings([{**ident(), "line": 3, "msg": "x"}]),
             baseline([entry()]), 0, None),
            ("degenerate: identity survives line churn", findings([ident(fingerprint="fp-1")]),
             baseline([entry(target_date="2026-10-05")]), 0, None),
        ])


class StaleIdentityLockIn(GateCase):
    """HR-15: eliminated identities must be locked in by shrinking the baseline."""
    def test_findings(self):
        self.table([
            ("eliminated identity still in baseline", findings([ident()]),
             baseline([entry(), entry(fingerprint="fp-2")]), 1, "HR-15"),
            ("everything eliminated", findings([ident(rule="kept")]),
             baseline([entry(), entry(rule="gone")]), 1, "HR-15"),
        ])

    def test_reports_eliminated_identity(self):
        code, out = self.run_gate(findings([ident()]),
                                  baseline([entry(), entry(fingerprint="fp-2")]))
        self.assertEqual(code, 1)
        self.assertIn("fp-2", out)
        self.assertIn("eliminated", out)


class TargetDate(GateCase):
    """Every baseline entry carries target and target_date; expired is a finding."""
    def test_findings(self):
        self.table([
            ("expired target date", findings([ident()]), baseline([entry(target_date="2026-10-03")]), 1, "HR-15"),
            ("long expired", findings([ident()]), baseline([entry(target_date="2025-01-01")]), 1, "HR-15"),
        ])

    def test_negatives(self):
        self.table([
            ("future target date", findings([ident()]), baseline([entry(target_date="2026-10-05")]), 0, None),
            ("today is not expired", findings([ident()]), baseline([entry(target_date=TODAY)]), 0, None),
            ("degenerate: expired on eliminated identity only", findings([ident(rule="kept")]),
             baseline([entry(rule="kept"), entry(rule="gone", target_date="2020-01-01")]), 1, None),
        ])


class LivenessAndZeroDebt(GateCase):
    """HR-3 liveness: exit 2 without a valid measurement proof; with it, zero
    debt is the achievable target state (empty findings + empty baseline = 0)."""
    def test_zero_debt_is_reachable(self):
        self.table([
            ("empty findings, empty baseline, proof valid", findings([]), baseline([]), 0, None),
            ("empty findings with entries locks them in", findings([]),
             baseline([entry(), entry(fingerprint="fp-2")]), 1, "HR-15"),
            ("proof valid with findings still measures", findings([ident()]),
             baseline([entry()]), 0, None),
        ])

    def test_no_proof_exits_two(self):
        code, out = self.run_gate({"findings": []}, baseline([entry()]))
        self.assertEqual(code, 2)
        self.assertIn("measured", out)


class Contract(GateCase):
    def test_missing_required_argument(self):
        with tempfile.TemporaryDirectory() as d:
            f = write_tmp(d, "f.json", findings([ident()]))
            out = subprocess.run([sys.executable, str(GATE), "--findings", str(f)],
                                 capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)
        self.assertIn("--baseline", out.stderr + out.stdout)

    def test_multiple_findings_aggregate(self):
        code, out = self.run_gate(findings([ident(), ident(fingerprint="fp-2")]),
                                  baseline([entry(), entry(fingerprint="fp-3", target_date="2020-01-01")]))
        self.assertEqual(code, 1)
        self.assertIn("fp-2", out)  # new
        self.assertIn("fp-3", out)  # expired

    def test_ok_reports_measurement(self):
        code, out = self.run_gate(findings([ident()]), baseline([entry()]))
        self.assertEqual(code, 0)
        self.assertIn("RATCHET OK", out)


class BaseComparison(GateCase):
    """HR-15 shrink-only: candidate baseline vs base-branch baseline. Growth
    requires a relaxation field (reason + reference) the base does not carry;
    shrink is always allowed. Without --base-baseline the comparison is
    declared NOT performed, never silently skipped."""
    def run_pair(self, candidate, base, find=None, extra=()):
        find = find if find is not None else findings([ident()])
        with tempfile.TemporaryDirectory() as d:
            f = write_tmp(d, "findings.json", find)
            b = write_tmp(d, "candidate.json", candidate)
            bb = write_tmp(d, "base.json", base)
            out = subprocess.run(
                [sys.executable, str(GATE), "--findings", str(f), "--baseline", str(b),
                 "--base-baseline", str(bb), "--today", TODAY, *extra],
                capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def test_growth_without_relaxation_fails(self):
        code, out = self.run_pair(
            baseline([entry(), entry(fingerprint="fp-2")]), baseline([entry()]))
        self.assertEqual(code, 1)
        self.assertIn("HR-15", out)

    def test_growth_with_valid_relaxation_passes(self):
        code, out = self.run_pair(
            baseline([entry(), entry(fingerprint="fp-2")],
                     relaxation={"reason": "new gate adopted", "reference": "repo#1", "actor": "maintainer"}),
            baseline([entry()]),
            find=findings([ident(), ident(fingerprint="fp-2")]))
        self.assertEqual(code, 0)

    def test_growth_with_malformed_relaxation_exits_two(self):
        code, out = self.run_pair(
            baseline([entry(), entry(fingerprint="fp-2")],
                     relaxation={"reason": "no reference"}),
            baseline([entry()]))
        self.assertEqual(code, 2)

    def test_base_relaxation_does_not_cover_new_growth(self):
        rel = {"reason": "old", "reference": "repo#0", "actor": "old-maintainer"}
        code, out = self.run_pair(
            baseline([entry(), entry(fingerprint="fp-2")], relaxation=rel),
            baseline([entry()], relaxation=rel))
        self.assertEqual(code, 1)

    def test_shrink_against_base_needs_no_relaxation(self):
        code, out = self.run_pair(
            baseline([entry()]),
            baseline([entry(), entry(fingerprint="fp-2")]))
        self.assertEqual(code, 0)
        self.assertIn("shrunk", out)

    def test_unreadable_base_baseline_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            f = write_tmp(d, "findings.json", findings([ident()]))
            b = write_tmp(d, "candidate.json", baseline([entry()]))
            bb = write_tmp(d, "base.json", "not json")
            out = subprocess.run(
                [sys.executable, str(GATE), "--findings", str(f), "--baseline", str(b),
                 "--base-baseline", str(bb), "--today", TODAY],
                capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)

    def test_without_base_baseline_declares_not_performed(self):
        code, out = self.run_gate(findings([ident()]), baseline([entry()]))
        self.assertEqual(code, 0)
        self.assertIn("NOT performed", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
