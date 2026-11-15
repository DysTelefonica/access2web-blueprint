"""The release preflight measures the tag object of the remote (issue #1541).

On a tag push, ``actions/checkout`` writes ``refs/tags/<tag>`` as a
lightweight tag on the commit and ``fetch-tags: true`` does not overwrite it.
A preflight that read the local ref rejected an annotated tag as lightweight
(the failure of cadete v1.3.2). These tests rebuild that runner state: a
remote holding the annotated tag and a clone whose tag is lightweight.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "release-preflight.sh"
TAG = "v1.0.0"
#: The bash first on PATH: on Windows, a bare "bash" resolves to WSL first.
BASH = shutil.which("bash")
GIT_ENV = {
    **os.environ,
    "GIT_AUTHOR_NAME": "t",
    "GIT_AUTHOR_EMAIL": "t@t",
    "GIT_COMMITTER_NAME": "t",
    "GIT_COMMITTER_EMAIL": "t@t",
}

pytestmark = pytest.mark.skipif(
    shutil.which("git") is None or BASH is None,
    reason="git and bash are needed",
)


def git(cwd: Path, *args: str) -> None:
    subprocess.run(["git", *args], cwd=cwd, env=GIT_ENV, check=True, capture_output=True)


def preflight(repo: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [BASH, str(SCRIPT), TAG],
        check=False,
        cwd=repo,
        env=GIT_ENV,
        capture_output=True,
        text=True,
    )


@pytest.fixture
def remote(tmp_path: Path) -> Path:
    """A bare remote whose ``main`` carries the annotated tag."""
    src = tmp_path / "src"
    git(tmp_path, "init", "-q", "-b", "main", str(src))
    git(src, "commit", "-q", "--allow-empty", "-m", "release")
    git(src, "tag", "-a", TAG, "-m", TAG)
    bare = tmp_path / "remote.git"
    git(tmp_path, "clone", "-q", "--bare", str(src), str(bare))
    return bare


def clone(remote: Path, target: Path) -> Path:
    git(remote.parent, "clone", "-q", str(remote), str(target))
    return target


def test_an_annotated_remote_tag_passes_although_the_local_one_is_lightweight(remote, tmp_path):
    runner = clone(remote, tmp_path / "runner")
    # What actions/checkout leaves on a tag push: the same name, lightweight.
    git(runner, "tag", "-f", TAG, f"{TAG}^{{commit}}")

    proc = preflight(runner)

    assert proc.returncode == 0, proc.stderr
    assert f"stable tag {TAG} verified" in proc.stdout


def test_a_lightweight_remote_tag_is_still_refused(tmp_path):
    src = tmp_path / "src"
    git(tmp_path, "init", "-q", "-b", "main", str(src))
    git(src, "commit", "-q", "--allow-empty", "-m", "release")
    git(src, "tag", TAG)
    bare = tmp_path / "remote.git"
    git(tmp_path, "clone", "-q", "--bare", str(src), str(bare))
    runner = clone(bare, tmp_path / "runner")

    proc = preflight(runner)

    assert proc.returncode == 1
    assert f"tag {TAG} is lightweight" in proc.stderr


def test_before_the_push_the_local_tag_is_measured(remote, tmp_path):
    runner = clone(remote, tmp_path / "runner")
    git(runner, "push", "-q", "origin", f":refs/tags/{TAG}")

    assert preflight(runner).returncode == 0

    git(runner, "tag", "-f", TAG, f"{TAG}^{{commit}}")
    proc = preflight(runner)
    assert proc.returncode == 1
    assert f"tag {TAG} is lightweight" in proc.stderr
