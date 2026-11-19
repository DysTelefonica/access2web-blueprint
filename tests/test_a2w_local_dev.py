"""Per-issue local stack of access2web (#1776): `skills/a2w-local-dev/assets/dev.sh`.

Checked through its public interface, the `env` subcommand and the usage errors, so no Docker
is needed: arguments, the per-issue project and ports, and the translation of a Windows
worktree `gitdir` when the script runs from WSL.
"""

import os
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "skills" / "a2w-local-dev" / "assets" / "dev.sh"
BASH = shutil.which("bash")

pytestmark = pytest.mark.skipif(BASH is None, reason="bash is required")


def run(*args: str, root: Path | None = None) -> subprocess.CompletedProcess[str]:
    env = {k: v for k, v in os.environ.items() if k not in ("GIT_DIR", "GIT_WORK_TREE")}
    if root is not None:
        env["A2W_ROOT"] = root.as_posix()
    return subprocess.run(
        [BASH, SCRIPT.as_posix(), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )


def env_of(issue: str, root: Path | None = None) -> dict[str, str]:
    done = run(issue, "env", root=root)
    assert done.returncode == 0, done.stderr
    return dict(line.split("=", 1) for line in done.stdout.splitlines() if "=" in line)


@pytest.mark.parametrize("args", [(), ("12",), ("doce", "env"), ("12", "deploy")])
def test_bad_arguments_print_the_usage_and_exit_2(args: tuple[str, ...]) -> None:
    done = run(*args)
    assert done.returncode == 2
    assert "uso: dev.sh <nº issue>" in done.stderr
    assert done.stdout == ""


def test_each_issue_gets_its_own_project_and_ports() -> None:
    a, b = env_of("1776"), env_of("1777")
    assert a["PROJECT"] == "a2w1776" and b["PROJECT"] == "a2w1777"
    assert (a["APP_PORT"], a["PG_PORT"]) == ("28776", "25776")
    assert (b["APP_PORT"], b["PG_PORT"]) == ("28777", "25777")


def test_a_plain_clone_needs_no_translation(tmp_path: Path) -> None:
    (tmp_path / ".git").mkdir()
    values = env_of("5", root=tmp_path)
    assert values["WINDOWS"] == "0"
    assert values["GIT_DIR"] == "" and values["GIT_WORK_TREE"] == ""


def test_a_windows_worktree_gitdir_is_translated_to_its_wsl_path(tmp_path: Path) -> None:
    (tmp_path / ".git").write_text(
        "gitdir: C:/00repos/codigo/access2web-blueprint/.git/worktrees/x\r\n", encoding="utf-8"
    )
    values = env_of("5", root=tmp_path)
    assert values["WINDOWS"] == "1"
    assert values["GIT_DIR"] == "/mnt/c/00repos/codigo/access2web-blueprint/.git/worktrees/x"
    assert values["GIT_WORK_TREE"] == tmp_path.as_posix()
