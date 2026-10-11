#!/usr/bin/env python3
# ci-pattern asset — ratchet gate: stable identities over a versioned baseline (#194)
"""Ratchet gate over normalized measuring-tool findings (HR-15 core, HR-16).

Input is a findings document produced by the measuring tool — an object
with a ``findings`` list of objects, each carrying the stable identity
``rule`` + ``path`` + ``fingerprint`` (content- or route-based, never a
mutable line number; extra keys from the tool are tolerated) — and a
versioned baseline (``baseline_version`` 1, ``entries`` with ``identity``,
``target`` and ``target_date``). The gate knows no tool: it consumes the
normalized output, so consumers keep their linters and this contract.

Rules (PR1 core):

- A finding whose identity is absent from the baseline is a finding (HR-15:
  the ratchet rejects what is new; only a baseline update can admit it).
- A baseline identity eliminated from the findings is reported and blocks:
  the baseline MUST be shrunk to lock the improvement in (HR-15 lock-in).
- A baseline entry whose ``target_date`` is past (strictly before ``--today``,
  ISO ``YYYY-MM-DD``) is a finding: the exemption lapsed.
- Liveness: a missing or unreadable baseline, a malformed document, or a
  findings document WITHOUT a valid measurement proof exits 2. The proof
  of life is data in the findings document itself — ``"measured": true``
  and ``"subjects"`` > 0, the count of subjects the tool actually
  measured. An empty ``findings`` list WITH a valid proof is the zero-debt
  target state: against an empty baseline it exits 0, and against a
  non-empty baseline it locks every entry in (exit 1). A gate that did
  not measure never says OK (HR-3).

- Shrink-only comparison against the base-branch baseline (HR-15): with
  ``--base-baseline`` the candidate baseline is compared against the base
  copy — every ADDED identity requires a ``relaxation`` field (``reason``
  and ``reference``, non-empty strings) the base copy does not already
  carry; SHRUNK identities are allowed and reported. Without
  ``--base-baseline`` the output declares the comparison NOT performed —
  never silently skipped.

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
FINDINGS_KEYS = frozenset({"measured", "subjects", "findings"})
BASELINE_KEYS = frozenset({"baseline_version", "entries"})
OPTIONAL_BASELINE_KEYS = frozenset({"relaxation"})
RELAXATION_KEYS = frozenset({"reason", "reference", "actor", "issue"})
RELAXATION_REQUIRED = frozenset({"reason", "reference"})
IDENTITY_KEYS = frozenset({"rule", "path", "fingerprint"})
ENTRY_KEYS = frozenset({"identity", "target", "target_date"})


class InvalidInput(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def identity_of(d: dict, label: str, where: str, strict: bool) -> tuple:
    """Extract (rule, path, fingerprint); strict=True requires exactly those keys,
    strict=False tolerates extra tool fields but requires the three identity fields."""
    if not isinstance(d, dict):
        raise InvalidInput(f"{label}: {where} is not an object")
    if strict and set(d) != IDENTITY_KEYS:
        raise InvalidInput(f"{label}: {where} identity must carry exactly {sorted(IDENTITY_KEYS)}")
    values = tuple(d.get(k) for k in ("rule", "path", "fingerprint"))
    if not all(isinstance(v, str) and v for v in values):
        raise InvalidInput(f"{label}: {where} identity has a missing, empty or non-string field")
    return values


def load_findings(path: str) -> list[tuple]:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InvalidInput(f"findings unreadable: {exc}") from exc
    if not isinstance(doc, dict) or set(doc) != FINDINGS_KEYS:
        raise InvalidInput(f"findings must carry exactly {sorted(FINDINGS_KEYS)}")
    if doc["measured"] is not True:
        raise InvalidInput('findings must declare "measured": true — the measuring tool '
                           "must attest a successful run; anything else is liveness exit 2")
    if isinstance(doc["subjects"], bool) or not isinstance(doc["subjects"], int) \
            or doc["subjects"] <= 0:
        raise InvalidInput('findings "subjects" must be an integer > 0 (the count of '
                           "measured subjects); zero or missing is liveness exit 2")
    raw = doc["findings"]
    if not isinstance(raw, list):
        raise InvalidInput("findings must be a list")
    identities = [identity_of(item, "findings", f"finding {i}", strict=False)
                  for i, item in enumerate(raw)]
    if len(set(identities)) != len(identities):
        raise InvalidInput("findings contain duplicate identities")
    return identities


def has_identity(exemption: dict) -> bool:
    """HR-33: a non-empty actor string or a positive issue integer."""
    actor = exemption.get("actor")
    issue = exemption.get("issue")
    return (isinstance(actor, str) and bool(actor.strip())) or (
        isinstance(issue, int) and not isinstance(issue, bool) and issue > 0)


def load_baseline(path: str) -> dict:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InvalidInput(f"baseline unreadable: {exc}") from exc
    if not isinstance(doc, dict) or not set(doc) <= (BASELINE_KEYS | OPTIONAL_BASELINE_KEYS) \
            or not BASELINE_KEYS <= set(doc):
        raise InvalidInput(f"baseline must carry {sorted(BASELINE_KEYS)} "
                           f"plus optionally {sorted(OPTIONAL_BASELINE_KEYS)}")
    relaxation = doc.get("relaxation")
    if "relaxation" in doc and (not isinstance(relaxation, dict)
                                or not set(relaxation) <= RELAXATION_KEYS
                                or not RELAXATION_REQUIRED <= set(relaxation)
                                or not all(isinstance(relaxation[k], str) and relaxation[k]
                                           for k in RELAXATION_REQUIRED)):
        raise InvalidInput('baseline "relaxation" must carry non-empty "reason" and '
                           '"reference" strings plus a verifiable identity — a '
                           'non-empty "actor" or a positive "issue"; a malformed field '
                           "never sits waiting to cover a later growth (HR-15, HR-33)")
    if "relaxation" in doc and not has_identity(relaxation):
        raise InvalidInput('baseline "relaxation" without verifiable identity (HR-33): '
                           'an exemption needs a non-empty "actor" or a positive "issue"')
    if doc["baseline_version"] != 1:
        raise InvalidInput(f"unsupported baseline_version {doc['baseline_version']!r}")
    entries = doc["entries"]
    if not isinstance(entries, list):
        raise InvalidInput("baseline entries must be a list")
    for i, entry in enumerate(entries):
        if not isinstance(entry, dict) or set(entry) != ENTRY_KEYS:
            raise InvalidInput(f"baseline entry {i} must carry exactly {sorted(ENTRY_KEYS)}")
        identity_of(entry["identity"], "baseline", f"entry {i}", strict=True)
        if not isinstance(entry["target"], str) or not entry["target"]:
            raise InvalidInput(f"baseline entry {i}: target must be a non-empty string")
        if not isinstance(entry["target_date"], str) or not ISO_DATE.fullmatch(entry["target_date"]):
            raise InvalidInput(f"baseline entry {i}: target_date must be ISO YYYY-MM-DD")
        try:
            date.fromisoformat(entry["target_date"])
        except ValueError as exc:
            raise InvalidInput(f"baseline entry {i}: invalid target_date: {exc}") from exc
    identities = [tuple(e["identity"][k] for k in ("rule", "path", "fingerprint")) for e in entries]
    if len(set(identities)) != len(identities):
        raise InvalidInput("baseline contains duplicate identities")
    return doc


def main(argv):
    ap = argparse.ArgumentParser(description="Ratchet gate (HR-15 core, HR-16).")
    ap.add_argument("--findings", required=True, help="normalized findings JSON from the measuring tool")
    ap.add_argument("--baseline", required=True, help="versioned baseline JSON")
    ap.add_argument("--today", default=None, help="ISO date for target_date expiry (default: today)")
    ap.add_argument("--base-baseline", default=None,
                    help="baseline copy on the base branch (shrink-only comparison, HR-15)")
    args = ap.parse_args(argv)
    today = date.fromisoformat(args.today) if args.today else date.today()
    try:
        findings = load_findings(args.findings)
        baseline_doc = load_baseline(args.baseline)
        baseline = baseline_doc["entries"]
        base_baseline = load_baseline(args.base_baseline) if args.base_baseline else None
    except InvalidInput as exc:
        print(f"RATCHET ERROR: {exc}", file=sys.stderr)
        return 2
    baseline_ids = {tuple(e["identity"][k] for k in ("rule", "path", "fingerprint")): e
                    for e in baseline}
    findings_ids = set(findings)
    report = []
    for ident in sorted(findings_ids - set(baseline_ids)):
        report.append(f"HR-15 new finding identity not in baseline "
                      f"(rule={ident[0]!r}, path={ident[1]!r}, fingerprint={ident[2]!r})")
    for ident in sorted(set(baseline_ids) - findings_ids):
        report.append(f"HR-15 identity eliminated from findings but still baselined — shrink "
                      f"the baseline to lock it in (rule={ident[0]!r}, path={ident[1]!r}, "
                      f"fingerprint={ident[2]!r})")
    for entry in baseline:
        ident = tuple(entry["identity"][k] for k in ("rule", "path", "fingerprint"))
        if ident in findings_ids and date.fromisoformat(entry["target_date"]) < today:
            report.append(f"HR-15 baseline target_date {entry['target_date']} expired "
                          f"(rule={ident[0]!r}, path={ident[1]!r}, "
                          f"fingerprint={ident[2]!r}; target={entry['target']!r})")
    if base_baseline is None:
        print("base comparison: NOT performed (--base-baseline absent; the caller "
              "supplies the base-branch copy when it exists)")
    else:
        base_ids = {tuple(e["identity"][k] for k in ("rule", "path", "fingerprint"))
                    for e in base_baseline["entries"]}
        added = baseline_ids.keys() - base_ids
        shrunk = base_ids - baseline_ids.keys()
        base_relaxation = base_baseline.get("relaxation")
        relaxation = baseline_doc.get("relaxation")
        covered = bool(added) and bool(relaxation) and relaxation != base_relaxation
        for ident in sorted(added):
            if covered:
                print(f"base comparison: growth of (rule={ident[0]!r}, path={ident[1]!r}, "
                      f"fingerprint={ident[2]!r}) is covered by the relaxation field "
                      f"(reason={relaxation['reason']!r}, reference={relaxation['reference']!r})")
            else:
                report.append(f"HR-15 baseline grows against the base branch with identity "
                              f"(rule={ident[0]!r}, path={ident[1]!r}, "
                              f"fingerprint={ident[2]!r}): uncovered — growth requires a "
                              "relaxation field with reason and reference the base copy "
                              "does not already carry")
        if shrunk:
            print(f"base comparison: performed — baseline shrunk by {len(shrunk)} "
                  "identity(ies) against the base branch (allowed without relaxation)")
    if report:
        print(f"RATCHET FAIL: {len(report)} finding(s)")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"RATCHET OK: {len(findings)} finding(s) measured against {len(baseline)} "
          "baseline entry(ies); no findings.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
