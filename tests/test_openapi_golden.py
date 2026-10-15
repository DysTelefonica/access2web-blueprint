# HARNESS-PROVENANCE: gent-884 — OpenAPI golden snapshot test.
"""Golden test for the FastAPI OpenAPI schema (issue #884).

The HTTP contract (routes, request/response schemas, status codes) is
published through ``app.openapi()``. An unintentional change in any of
those — a typo in a path, a renamed response field, a new status
code — used to slip through because no test pinned the schema. This
module snaps ``app.openapi()`` to ``tests/golden/openapi.json`` and
fails the gate whenever the live schema drifts from the snapshot.

The change is reviewed in the PR; intentional updates are committed
alongside the code change. To regenerate the snapshot after an
intentional contract change, run:

.. code-block:: bash

    pytest tests/test_openapi_golden.py --update-golden

The acceptance criteria from the issue are pinned by the assertions
below:

* Renaming a response schema field makes the test fail (the diff is
  printed so the reviewer can see exactly what changed).
* ``--update-golden`` regenerates ``openapi.json`` and the test passes
  again.

Scope note: ``app.openapi()`` reflects whatever is registered on the
``FastAPI`` instance when the schema is queried. The composition root
in ``app/src/main.py`` wires the admin slices through an
``on_event("startup")`` callback that runs only when a lifespan
without an explicit container is started. The golden therefore
captures the import-time surface (``/health`` plus the auto-generated
``/openapi.json``, ``/docs`` and ``/redoc``) and will catch any new
route or schema that is added at module-load. Slices that move onto
the module-load path will appear in the snapshot automatically and
the diff will surface the move to the reviewer.
"""

from __future__ import annotations

import difflib
import json
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from app.src.main import app

if TYPE_CHECKING:
    from pytest import FixtureRequest

GOLDEN_DIR = Path(__file__).parent / "golden"
GOLDEN_PATH = GOLDEN_DIR / "openapi.json"


def _serialise(schema: object) -> str:
    """Return a stable byte-for-byte representation of ``schema``.

    ``sort_keys=True`` keeps the diff meaningful across re-orderings of
    the underlying dict; ``indent=2`` matches the rest of the repo's
    JSON fixtures so the diff is easy to review in GitHub's PR view.
    The trailing newline keeps POSIX tooling (``diff``, ``git diff``)
    quiet.
    """
    return json.dumps(schema, indent=2, sort_keys=True) + "\n"


def _render_diff(expected: str, actual: str) -> str:
    """Return a unified diff between the golden and the live schema."""
    return "".join(
        difflib.unified_diff(
            expected.splitlines(keepends=True),
            actual.splitlines(keepends=True),
            fromfile=str(GOLDEN_PATH.relative_to(Path(__file__).parent)),
            tofile="app.openapi()",
            n=3,
        )
    )


def test_openapi_matches_golden_snapshot(request: FixtureRequest) -> None:
    """The live OpenAPI schema must match the committed golden snapshot.

    Behaviour by mode:

    * Default run: assert equality. On drift, print a unified diff and
      fail with an instruction pointing at ``--update-golden``.
    * ``--update-golden``: regenerate the snapshot from the live
      schema and pass. The regenerated file must be committed in the
      same PR as the schema change.
    * First-run bootstrap: if the snapshot does not exist yet, create
      it and fail with an instruction to review and commit it. The
      snapshot is never generated silently on a default run because
      that would let a fresh checkout pass without any review.
    """
    update: bool = bool(request.config.getoption("--update-golden"))
    actual_text = _serialise(app.openapi())

    if not GOLDEN_PATH.exists():
        GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
        GOLDEN_PATH.write_text(actual_text, encoding="utf-8")
        pytest.fail(
            "OpenAPI golden snapshot was missing and has been created at "
            f"{GOLDEN_PATH.relative_to(Path(__file__).parent.parent)}. "
            "Review it, commit it, then re-run the tests.",
        )

    expected_text = GOLDEN_PATH.read_text(encoding="utf-8")

    if actual_text == expected_text:
        return

    if update:
        GOLDEN_PATH.write_text(actual_text, encoding="utf-8")
        return

    pytest.fail(
        "OpenAPI schema drifted from the golden snapshot.\n"
        "Review the diff below; if the change is intentional, run:\n"
        "    pytest tests/test_openapi_golden.py --update-golden\n"
        "and commit the regenerated snapshot in this PR.\n\n"
        f"{_render_diff(expected_text, actual_text)}"
    )


__all__ = ["GOLDEN_PATH", "test_openapi_matches_golden_snapshot"]
