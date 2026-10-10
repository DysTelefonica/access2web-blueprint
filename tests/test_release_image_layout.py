"""The release image builds and its CMD resolves the package it ships.

Issue #772: the committed `app/Dockerfile` could not produce an image at all —
`pip install .` runs before the README the metadata reads is in the layer, and
the runtime copied the source to a path where `app.src.main` does not resolve.
Both failure modes are invisible to the test suite because they only appear when
Docker builds the image, so this test pins the three things the build needs.
"""

from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
DOCKERFILE = REPO_ROOT / "app" / "Dockerfile"
DOCKERFILE_TEXT = DOCKERFILE.read_text(encoding="utf-8")

INSTALL = "pip install --no-cache-dir ."


def test_metadata_inputs_are_in_the_layer_before_the_install() -> None:
    """`readme` and the wheel's source must exist when hatchling reads them."""
    before_install = DOCKERFILE_TEXT.split(INSTALL, 1)[0]
    assert "COPY pyproject.toml README.md ./" in before_install, (
        "README.md is not in the layer before `pip install .`: pyproject.toml declares "
        'readme = "README.md" and the build dies with "Readme file does not exist"'
    )
    assert "COPY src/ ./src/" in before_install, (
        "src/ is not in the layer before `pip install .`: the wheel only-includes src, "
        "so hatchling has nothing to ship"
    )


def test_runtime_exposes_the_package_the_cmd_imports() -> None:
    """`uvicorn app.src.main:app` needs an importable `app` package on sys.path."""
    assert "app.src.main:app" in DOCKERFILE_TEXT, "the CMD no longer names the package under test"
    assert "COPY src/ /app/app/src/" in DOCKERFILE_TEXT, (
        "the runtime copies the source somewhere `app.src` does not resolve: the application "
        "imports itself as `app.src...` in every module"
    )
    assert 'PYTHONPATH="/app"' in DOCKERFILE_TEXT, (
        "PYTHONPATH does not include /app, so `app.src.main` is not importable at runtime"
    )
