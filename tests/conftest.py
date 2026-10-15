# HARNESS-PROVENANCE: gent-884 — ``--update-golden`` flag for OpenAPI golden test.
"""Root-level pytest configuration for the access2web-blueprint test suite.

Issue #884 introduces an OpenAPI golden snapshot test
(``tests/test_openapi_golden.py``) that needs a CLI flag to regenerate
the snapshot on purpose. The flag is registered here so the test can
read it via ``request.config.getoption("--update-golden")`` without
owning the lifecycle of the option.

The flag is intentionally opt-in (``store_true``): the default test
run must fail when the schema drifts so reviewers cannot accidentally
ship a contract change without noticing. Regenerating the snapshot is
an explicit, single-step operation:

.. code-block:: bash

    pytest tests/test_openapi_golden.py --update-golden

The committed snapshot then rides in the same PR as the schema change
so reviewers can see the diff side-by-side.
"""

from __future__ import annotations

import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    """Register the ``--update-golden`` CLI flag."""
    parser.addoption(
        "--update-golden",
        action="store_true",
        default=False,
        help=(
            "Regenerate the OpenAPI golden snapshot "
            "(tests/golden/openapi.json) instead of comparing against it. "
            "Use after an intentional schema change; the new snapshot "
            "must be committed in the same PR."
        ),
    )


__all__ = ["pytest_addoption"]
