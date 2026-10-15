# HARNESS-PROVENANCE: gent-884 — ``--update-golden`` flag for OpenAPI golden test.
# HARNESS-PROVENANCE: gent-883 — pytest-socket + env-isolation autouse fixture.
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

Issue #883 adds the unit-suite isolation contract:

  * Unit tests cannot reach the public network (pytest-socket's
    ``--disable-socket --allow-unix-socket --allow-hosts=127.0.0.1``
    wired in ``app/pyproject.toml::[tool.pytest.ini_options].addopts``).
  * Unit tests cannot read the developer's ``.env``: the
    ``_isolate_unit_env`` autouse fixture below wipes any var outside
    ``ENV_ALLOWLIST_EXACT`` or matching a prefix in
    ``ENV_ALLOWLIST_PREFIXES`` before the test body runs.
  * Integration (``@pytest.mark.integration``) and end-to-end
    (``@pytest.mark.e2e``) tests opt out of both: they need real
    database URLs, cloud credentials, and staging tokens. The marker
    check happens before any cleanup so the fixture is a no-op for them.
"""

from __future__ import annotations

import os

import pytest

# Env vars unit tests may rely on. Anything outside this set is wiped by
# ``_isolate_unit_env`` before the test body runs. Prefixes are matched
# with ``str.startswith`` so ``PYTEST_*``, ``GITHUB_*``, ``RUNNER_*`` and
# the like stay available to the suite without becoming a wide-open door.
ENV_ALLOWLIST_EXACT: frozenset[str] = frozenset(
    {
        # System / shell
        "PATH",
        "HOME",
        "USER",
        "SHELL",
        "TERM",
        "LANG",
        "OS",
        "HOSTNAME",
        "TMPDIR",
        "TEMP",
        "TMP",
        "LSCOLORS",
        "EDITOR",
        "VISUAL",
        # Python interpreter
        "PYTHONPATH",
        "PYTHONHOME",
        "PYTHONIOENCODING",
        "PYTHONDONTWRITEBYTECODE",
        # Process / virtual env (CI runners use these to find the interpreter)
        "VIRTUAL_ENV",
        "VIRTUAL_ENV_PROMPT",
        "CONDA_DEFAULT_ENV",
        "CONDA_PREFIX",
        "PIP_CONFIG_FILE",
        "PIP_DISABLE_PIP_VERSION_CHECK",
        # CI signals (necessary so the suite knows it's running in CI)
        "CI",
        "CONTINUOUS_INTEGRATION",
    }
)
ENV_ALLOWLIST_PREFIXES: tuple[str, ...] = (
    "LC_",
    "PYTHON",
    "CONDA_",
    "PIP_",
    "GITHUB_",
    "RUNNER_",
    "ACTIONS_",
)


def _is_env_allowed(key: str) -> bool:
    """Return True if ``key`` is on the env allowlist (exact or prefix)."""
    if key in ENV_ALLOWLIST_EXACT:
        return True
    return any(key.startswith(prefix) for prefix in ENV_ALLOWLIST_PREFIXES)


@pytest.fixture(autouse=True)
def _isolate_unit_env(monkeypatch: pytest.MonkeyPatch, request: pytest.FixtureRequest) -> None:
    """Wipe env vars outside the allowlist (issue #883).

    Integration (``@pytest.mark.integration``) and end-to-end
    (``@pytest.mark.e2e``) tests opt out: they need real database URLs,
    cloud credentials, and staging tokens that do not belong in a unit
    suite. The marker check happens before any cleanup so the fixture
    is a no-op for those tests.
    """
    if "integration" in request.keywords or "e2e" in request.keywords:
        return
    for key in list(os.environ):
        if not _is_env_allowed(key):
            monkeypatch.delenv(key, raising=False)


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
