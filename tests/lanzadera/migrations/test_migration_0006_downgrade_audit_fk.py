# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #924
"""Regression test for issue #924: 0006 downgrade leaves audit rows with actor_id
pointing at users, blocking the 0005 -> 0004 downgrade with FK violation.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — the downgrade does not null out ``actor_id`` on
  audit before letting 0004 delete users; the test fails.
  GREEN once the fix lands — ``UPDATE {SCHEMA}.audit SET actor_id = NULL``
  appears before the DELETE; the test passes.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest


def _migration_0006_downgrade_body() -> str:
    path = (
        Path(__file__).resolve().parents[3]
        / "app"
        / "migrations"
        / "versions"
        / "0006_seed_audit.py"
    )
    module = ast.parse(path.read_text())
    for node in module.body:
        if isinstance(node, ast.FunctionDef) and node.name == "downgrade":
            return ast.unparse(node)
    raise AssertionError("0006_seed_audit.py must define a downgrade() function")


def test_0006_downgrade_nulls_actor_id_before_delete() -> None:
    body = _migration_0006_downgrade_body()
    assert "UPDATE" in body, "0006 downgrade must UPDATE audit.actor_id before DELETE"
    assert "actor_id = NULL" in body, (
        "0006 downgrade must set audit.actor_id to NULL"
    )
    assert "DELETE" in body, "0006 downgrade must DELETE the seeded audit rows"
    update_pos = body.find("UPDATE")
    delete_pos = body.find("DELETE")
    assert update_pos < delete_pos, (
        "0006 downgrade must null out audit.actor_id BEFORE the DELETE so 0004 "
        "can drop users without the FK fk_audit_actor_id_users"
    )
