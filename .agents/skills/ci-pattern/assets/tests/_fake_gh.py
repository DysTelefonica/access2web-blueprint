"""The fake ``gh`` of the asset tests, reached by explicit path (#1609).

Shared by the asset test suites (import it via ``sys.path`` from ``assets/tests``).
An asset that runs ``gh`` takes ``CI_PATTERN_GH`` when set (a ``.py`` path runs
with the asset's own Python) and ``gh`` from the PATH otherwise. The tests set
``CI_PATTERN_GH`` to the absolute path of the fake's Python script, so no test
depends on how the PATH resolves: on Windows an extensionless ``gh`` is never
found and ``gh.exe``, the real one, wins; a ``gh.cmd`` goes through ``cmd.exe``,
which splits an argument such as ``?per_page=100&page=1`` at the ``&``.
A ``gh`` launcher (``gh.cmd`` on Windows) sits next to the script for the tests
that prove the PATH fallback; those run on POSIX only.

``run_guarded`` and ``guard_fake_gh`` fail closed before any process starts when
the ``gh`` the asset would run lies outside the test's fixture dir. On Windows
they also refuse without ``CI_PATTERN_GH``: an asset that runs ``gh`` by name
lets ``CreateProcess`` search the PATH for ``gh.exe`` only, so a fake found by
``shutil.which`` proves nothing there.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

GH_ENV = "CI_PATTERN_GH"
PATH_TESTS_SKIP = "#1609: on Windows the fake gh is reached only through CI_PATTERN_GH"


class RealGhRefusedError(AssertionError):
    """The asset would run a ``gh`` outside the fixture dir: the test does not run."""


def install_fake_gh(bin_dir: Path, stub: str) -> Path:
    """Write ``stub`` (Python source) and a ``gh`` launcher; return the script's absolute path."""
    bin_dir = Path(bin_dir).resolve()
    bin_dir.mkdir(parents=True, exist_ok=True)
    script = bin_dir / "gh_stub.py"
    script.write_text(stub, encoding="utf-8")
    if os.name == "nt":
        (bin_dir / "gh.cmd").write_text(f'@"{sys.executable}" "{script}" %*\r\n', encoding="utf-8")
    else:
        launcher = bin_dir / "gh"
        launcher.write_text(f'#!/bin/sh\nexec "{sys.executable}" "{script}" "$@"\n', encoding="utf-8")
        launcher.chmod(0o755)
    return script


def resolve_gh(env: dict) -> str | None:
    """The ``gh`` an asset would run under ``env``: the same rule the assets apply."""
    explicit = env.get(GH_ENV)
    if explicit and explicit.endswith(".py"):
        return explicit if Path(explicit).is_file() else None
    return shutil.which(explicit or "gh", path=env.get("PATH", os.defpath))


def guard_fake_gh(env: dict, fixture_dir: Path | str) -> None:
    """Raise ``RealGhRefusedError`` unless the resolved ``gh`` lies inside ``fixture_dir``."""
    resolved = resolve_gh(env)
    root = Path(fixture_dir).resolve()
    # On Windows CreateProcess resolves a bare name to gh.exe only, so anything but an
    # absolute path in CI_PATTERN_GH could still reach the real gh (#1609).
    explicit = env.get(GH_ENV)
    by_name_on_windows = os.name == "nt" and not (explicit and Path(explicit).is_absolute())
    if resolved is None or by_name_on_windows or not Path(resolved).resolve().is_relative_to(root):
        msg = (f"refused: the gh this test would run ({resolved}) is outside the fixture dir "
               f"{root}; set {GH_ENV} to the fake gh. The test did not run and no call was made (#1609)")
        raise RealGhRefusedError(msg)


def run_guarded(cmd: list[str], *, env: dict, fixture_dir: Path | str,
                **kwargs: object) -> subprocess.CompletedProcess:
    """``subprocess.run(cmd, env=env, ...)`` only after ``guard_fake_gh`` passes."""
    guard_fake_gh(env, fixture_dir)
    return subprocess.run(cmd, env=env, check=False, **kwargs)  # noqa: S603 - fixed argv, no shell
