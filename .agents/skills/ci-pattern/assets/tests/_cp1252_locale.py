"""Run a script or a snippet under a forced cp1252 locale on any host (#1596, #1601).

Shared by the asset test suites (import it via ``sys.path`` from ``assets/tests``).
UTF-8 mode is switched off with ``-X utf8=0`` (it overrides PYTHONUTF8) and the
locale lookups return cp1252, so the child decodes like a default Windows host
even on Linux CI. Snippets print with ``ascii()`` so the parent reads the
result whatever the child's stdout encoding is.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

CP1252_PATCH = (
    "import locale\n"
    "locale.getencoding = lambda: 'cp1252'\n"
    "locale.getpreferredencoding = lambda do_setlocale=True: 'cp1252'\n"
)
RUN_SCRIPT = (
    "import runpy, sys\n"
    "sys.argv = sys.argv[1:]\n"
    "runpy.run_path(sys.argv[0], run_name='__main__')\n"
)


def _cp1252_child(argv: list[str], cwd: Path | None, env: dict | None = None,
                  timeout: float | None = None) -> tuple[int, str, str]:
    env = {k: v for k, v in {**os.environ, **(env or {})}.items()
           if k not in ("PYTHONUTF8", "PYTHONIOENCODING")}
    done = subprocess.run(  # noqa: S603 - fixed argv, no shell
        [sys.executable, "-B", "-X", "utf8=0", *argv],
        capture_output=True,
        env=env,
        cwd=cwd,
        timeout=timeout,
        check=False,
    )
    return (done.returncode, done.stdout.decode("utf-8", errors="replace"),
            done.stderr.decode("utf-8", errors="replace"))


def run_under_cp1252_locale(script: Path, *args: str, cwd: Path | None = None,
                            env: dict | None = None, timeout: float | None = None) -> tuple[int, str, str]:
    """Run ``script`` as ``__main__`` with ``args`` in a cp1252-locale child.

    ``env`` is merged over ``os.environ``; ``timeout`` is passed to ``subprocess.run``.
    """
    return _cp1252_child(["-c", CP1252_PATCH + RUN_SCRIPT, str(script), *args], cwd, env, timeout)


def run_code_under_cp1252_locale(code: str, cwd: Path | None = None) -> tuple[int, str, str]:
    """Run the Python ``code`` snippet in a cp1252-locale child."""
    return _cp1252_child(["-c", CP1252_PATCH + code], cwd)
