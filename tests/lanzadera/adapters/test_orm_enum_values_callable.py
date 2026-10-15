# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #931
"""Regression test for issue #931: sa.Enum must use values_callable.

The three ``sa.Enum(...)`` definitions in the persistence layer map Python
``StrEnum`` members to Postgres ENUM labels. Without ``values_callable``,
SQLAlchemy uses the member NAMES (uppercase, e.g. ``PASSWORD_RESET_REQUIRED``)
as labels, which do not match the lowercase labels in Postgres (e.g.
``password_reset_required``). asyncpg rejects the INSERT with
``InvalidTextRepresentationError``.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — at least one ``sa.Enum`` lacks
  ``values_callable``; the test fails.
  GREEN once the fix lands — every ``sa.Enum`` in the adapter has
  ``values_callable``; the test passes.
"""

from __future__ import annotations

import ast
from pathlib import Path
from typing import cast

# Tables that must map Python enum VALUES (lowercase) to Postgres ENUM
# labels. Adding a new table here forces the test to cover it.
TABLES_WITH_ENUM = (
    "app/src/modules/lanzadera/adapters/persistence/repositories/user_table.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/app_table.py",
)


def _find_sa_enum_calls(source: str) -> list[ast.Call]:
    """Return every ``sa.Enum(...)`` call in the file, as AST nodes."""
    tree = ast.parse(source)
    found: list[ast.Call] = []

    class Visitor(ast.NodeVisitor):
        def visit_Call(self, node: ast.Call) -> None:
            func = node.func
            # ``sa.Enum(...)`` parses as Attribute(value=Name('sa'), attr='Enum')
            if (
                isinstance(func, ast.Attribute)
                and func.attr == "Enum"
                and isinstance(func.value, ast.Name)
                and func.value.id == "sa"
            ):
                found.append(node)
            self.generic_visit(node)

    Visitor().visit(tree)
    return found


def _call_uses_values_callable(call: ast.Call) -> bool:
    return any(kw.arg == "values_callable" for kw in call.keywords)


def _call_source_segment(source_lines: list[str], node: ast.AST) -> str:
    typed = cast(ast.expr, node)
    start = typed.lineno - 1
    end = getattr(typed, "end_lineno", typed.lineno)
    return "\n".join(source_lines[start:end])


def test_orm_tables_use_values_callable_on_sa_enum() -> None:
    """Every ``sa.Enum(...)`` in the listed tables must pass ``values_callable``.

    ``sa.Enum(UserStatus, name="user_status")`` (no ``values_callable``) is
    a bug: it persists ``PASSWORD_RESET_REQUIRED`` instead of
    ``password_reset_required``. Issue #931.
    """
    for relative in TABLES_WITH_ENUM:
        path = Path(__file__).resolve().parents[3] / relative
        source = path.read_text()
        calls = _find_sa_enum_calls(source)
        assert calls, f"{relative} has no sa.Enum at all"
        source_lines = source.splitlines()
        for call in calls:
            segment = _call_source_segment(source_lines, call)
            assert _call_uses_values_callable(call), (
                f"{relative} declares sa.Enum without values_callable:\n"
                f"{segment}\n"
                f"Without it SQLAlchemy sends Python enum member NAMES "
                f"(uppercase) as the Postgres ENUM label, which does not "
                f"match the lowercase label in 0001. See #931."
            )
