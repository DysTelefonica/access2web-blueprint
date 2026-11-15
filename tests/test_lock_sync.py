"""#1455: the audited lock and the lock the repository declares must be one resolution.

``app/uv.lock`` is the resolution of the project. ``app/requirements-dev.lock`` is what
``uv export`` renders from it, with the hash of every wheel — and that exported file is
what the CI installs (``pip install --require-hashes --requirement app/requirements-dev.lock``)
and what ``pip-audit`` audits.

If the two drift, the gate measures a set of packages the repository no longer declares,
while ``.github/required-jobs-policy.json`` keeps promising a coverage it does not have:
its ``state_dependent_gates`` entry for ``pip-audit`` declares ``app/uv.lock`` among the
manifests the gate covers, and the lane fires when that file changes — but the scan reads
the other lock. The two files agree today by convention; this makes the agreement a
measured fact instead.

``uv export`` runs with ``--no-emit-project``, so the project itself (an editable source,
the working tree rather than a wheel from an index) is not part of the exported lock and is
not a divergence either.

The checks take the two files as text, so the red fixtures below feed doctored text to the
same functions that judge the real ones.
"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
UV_LOCK = ROOT / "app" / "uv.lock"
REQUIREMENTS_LOCK = ROOT / "app" / "requirements-dev.lock"

#: One line per package: ``name==version --hash=... --hash=...``.
REQUIREMENT = re.compile(r"^([A-Za-z0-9_.\-]+)==([^\s\\]+)", re.MULTILINE)


def uv_resolution(text: str) -> tuple[dict[str, str], set[str]]:
    """(package -> version, local packages) of the declared resolution.

    A local package has no wheel in an index (its source is the working tree), so it is
    not part of the exported lock and not a divergence.
    """
    document = tomllib.loads(text)
    declared: dict[str, str] = {}
    local: set[str] = set()
    for package in document["package"]:
        declared[package["name"]] = package["version"]
        source = package.get("source", {})
        if "editable" in source or "virtual" in source:
            local.add(package["name"])
    return declared, local


def exported_requirements(text: str) -> dict[str, str]:
    """package -> version of the exported, hash-pinned lock."""
    return {match.group(1): match.group(2) for match in REQUIREMENT.finditer(text)}


def divergence(declared: dict[str, str], local: set[str], exported: dict[str, str]) -> list[str]:
    """Every way the audited lock and the declared resolution disagree."""
    expected = {name: version for name, version in declared.items() if name not in local}
    problems: list[str] = []
    for name in sorted(set(expected) - set(exported)):
        problems.append(f"{name} {expected[name]} is in app/uv.lock and not in the audited lock")
    for name in sorted(set(exported) - set(expected)):
        problems.append(f"{name} {exported[name]} is in the audited lock and not in app/uv.lock")
    for name in sorted(set(expected) & set(exported)):
        if expected[name] != exported[name]:
            problems.append(
                f"{name}: app/uv.lock has {expected[name]}, the audited lock has {exported[name]}"
            )
    return problems


def test_the_audited_lock_is_the_resolution_the_repository_declares() -> None:
    """The lock `pip-audit` audits is the resolution `app/uv.lock` declares.

    A divergence here means the gate measures something else from what the repository
    declares, and the `pip-audit` entry of the required-jobs policy promises a coverage
    (`app/uv.lock`) that nothing measures.
    """
    declared, local = uv_resolution(UV_LOCK.read_text(encoding="utf-8"))
    exported = exported_requirements(REQUIREMENTS_LOCK.read_text(encoding="utf-8"))
    problems = divergence(declared, local, exported)
    assert not problems, (
        "app/requirements-dev.lock is not the export of app/uv.lock; regenerate it with\n"
        "  uv export --project app --frozen --no-emit-project --extra dev \\\n"
        "    --format requirements-txt --no-annotate --no-header\n"
        "so `pip-audit` audits the resolution the repository declares:\n  " + "\n  ".join(problems)
    )


def test_the_audited_lock_carries_the_hash_of_every_package() -> None:
    """The audited lock is installed and audited with `--require-hashes`.

    Without a hash per package the install is not reproducible and the audit measures a
    set the CI could have resolved differently.
    """
    lines = [
        line
        for line in REQUIREMENTS_LOCK.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    ]
    packages = exported_requirements("\n".join(lines))
    assert packages, "the audited lock is empty"
    unhashed = [name for name in packages if f"{name}=={packages[name]} " not in "\n".join(lines)]
    assert not unhashed, f"packages without a hash in the audited lock: {unhashed}"


def test_the_project_itself_is_not_exported() -> None:
    """`uv export --no-emit-project` omits the editable package: it is not a wheel.

    A local package is the working tree, so it has no version to compare and its absence
    from the audited lock is not a divergence.
    """
    declared, local = uv_resolution(UV_LOCK.read_text(encoding="utf-8"))
    exported = exported_requirements(REQUIREMENTS_LOCK.read_text(encoding="utf-8"))
    assert local, "app/uv.lock must declare the project as an editable package"
    assert not (local & set(exported)), (
        "the project itself is not a wheel from an index: it must not be in the audited lock"
    )
    assert not divergence(declared, local, exported)


def test_red_a_package_missing_from_the_audited_lock_is_a_divergence() -> None:
    """A package the repository declares and the gate does not measure is a finding."""
    problems = divergence({"alpha": "1.0", "beta": "2.0"}, set(), {"alpha": "1.0"})
    assert problems == ["beta 2.0 is in app/uv.lock and not in the audited lock"]


def test_red_a_package_only_in_the_audited_lock_is_a_divergence() -> None:
    """A package the gate measures and the repository does not declare is a finding."""
    problems = divergence({"alpha": "1.0"}, set(), {"alpha": "1.0", "gamma": "3.0"})
    assert problems == ["gamma 3.0 is in the audited lock and not in app/uv.lock"]


def test_red_a_different_version_is_a_divergence() -> None:
    """The same package at two versions: the gate measures a resolution nobody declared."""
    problems = divergence({"alpha": "1.0"}, set(), {"alpha": "1.1"})
    assert problems == ["alpha: app/uv.lock has 1.0, the audited lock has 1.1"]


def test_red_a_local_package_is_not_a_divergence() -> None:
    """The project itself is omitted by `--no-emit-project`: its absence is not a finding."""
    declared = {"app-lanzadera-mvp": "0.0.0+phase0", "alpha": "1.0"}
    problems = divergence(declared, {"app-lanzadera-mvp"}, {"alpha": "1.0"})
    assert not problems
