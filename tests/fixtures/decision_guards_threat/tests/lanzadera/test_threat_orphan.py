# HARNESS-PROVENANCE: deterministic-quality-harness v1.6
# + architectural-guards-over-metrics DG-99 — test_threat_orphan.py
"""Threat fixture for the path-collision vector (slice 5, DG-5/DG-10).

The fake ID ``DG-99`` would trigger ``orphan_guard`` if the production
traversal (``--root .``) ever read this file. It MUST be invisible because
``EXCLUDED_PARTS`` in ``scripts/check_decision_guards.py`` contains
``"fixtures"``; that exclusion is the safety net this fixture documents.
"""