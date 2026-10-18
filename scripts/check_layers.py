#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 — assets/scripts/check_layers.py
"""Hexagonal layer gate: dependency direction, vertical slicing, and layer purity.

Stdlib only, so the gate runs before the project has installed anything. Walks the AST of every
module under the root package, resolves each import to a ``(module, layer)`` pair, and fails on
any edge the architecture does not allow.

Regex cannot do this job: it breaks on line continuations, comments, aliased imports, and relative
imports. Hard Rule 9 requires the AST.

RETIRING (#1013): the layer rules are enforced by the docs-architecture job from the
``architecture`` block of docs/architecture.md and .importlinter. This module no longer has an
entry point or a baseline; it only keeps the layer constants that tests/test_architecture_doc.py
reads, and is deleted by the next PR of the chain.
"""

from __future__ import annotations

import ast
from dataclasses import dataclass
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION — adjust this block when instantiating. Everything below it is mechanism.
# --------------------------------------------------------------------------------------------

ROOT_PACKAGE = "app"

#: Path segments between the root package and the module name. This repository lays modules out
#: as `app/src/modules/<module>/<layer>/`, not `app/<module>/<layer>/`, so without this the
#: classifier matched nothing: it reported 21 of 22 files unclassified and called each one a
#: violation. A gate that cannot place the code it walks has no verdict to give (Hard Rule 18) —
#: and 21 "violations" that are really 21 shrugs is the most misleading number it could print.
MODULE_PREFIX: tuple[str, ...] = ("src", "modules")

#: Modules every other module may depend on. Keep this set as small as it can possibly be; each
#: entry is a hole in the vertical slicing rule.
CROSS_CUTTING_MODULES = frozenset({"core"})

#: Which layers a given layer may import from. A layer absent from this table is not a layer.
ALLOWED_IMPORTS: dict[str, frozenset[str]] = {
    "domain": frozenset({"domain"}),
    "ports": frozenset({"domain", "ports"}),
    "application": frozenset({"domain", "ports", "application"}),
    "adapters": frozenset({"domain", "ports", "adapters", "shared"}),
    "shared": frozenset({"domain", "ports", "shared"}),
    "delivery": frozenset({"domain", "ports", "application", "adapters", "shared", "delivery"}),
    "di": frozenset({"domain", "ports", "application", "adapters", "shared", "delivery", "di"}),
    # M01..M04 of #225/#257 (#225=#225 baseline) land read-only
    # extractors, staging writers, and promotion logic under
    # ``app/src/modules/expedientes/migration/``. The layer is a sibling
    # of ``adapters`` (it touches the same cross-cutting infrastructure)
    # but the only thing it imports from the rest of the module is the
    # domain's value objects. M02..M04 do not yet exist; the layer is
    # here so M01's ``extractor.py`` and ``watermark.py`` pass the
    # ``check_layers.py`` ratchet without being treated as
    # "unclassified".
    "migration": frozenset({"domain", "ports", "adapters", "shared", "migration"}),
}

#: Layers that must not touch a framework at all.
PURE_LAYERS = frozenset({"domain", "ports", "application"})

#: Top-level distributions banned inside PURE_LAYERS.
FORBIDDEN_IN_PURE_LAYERS = frozenset(
    {
        "fastapi",
        "sqlalchemy",
        "alembic",
        "jinja2",
        "httpx",
        "asyncpg",
        "starlette",
        "pydantic_settings",
    }
)

#: Paths excluded from the walk. Hard Rule 14: this is the declared untestable/ungoverned
#: boundary. Keep it thin and keep it honest — every entry here is code nobody is checking.
EXCLUDED_PARTS = frozenset(
    {"__pycache__", ".venv", "venv", "build", "dist", "migrations", "pytest_plugin"}
)

#: Files that sit outside the module/layer structure ON PURPOSE, each with the reason.
#:
#: Everything else that cannot be classified is reported as unclassified and fails, because
#: "I could not place this file" and "this file is clean" must never share a verdict
#: (Hard Rule 18). This list is what separates a deliberate exception from a blind spot, and
#: it is deliberately short: every entry is a file the gate does not check.
UNLAYERED_FILES: dict[str, str] = {
    "app/src/main.py": "composition root — the one place allowed to wire every layer together",
}


# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class Violation:
    key: str
    detail: str
    file: str
    line: int


def _iter_source_files(root: Path) -> list[Path]:
    package_root = root / ROOT_PACKAGE
    if not package_root.is_dir():
        return []
    files = []
    for path in sorted(package_root.rglob("*.py")):
        if EXCLUDED_PARTS.intersection(path.parts):
            continue
        files.append(path)
    return files


def _is_package_marker(path: Path, root: Path) -> bool:
    """True for an ``__init__.py`` above the module level.

    Such a file cannot carry a layer: there is no ``<module>/<layer>/`` prefix for it to sit
    under. An ``__init__.py`` INSIDE a layer classifies normally and is checked like any other
    file, so this exempts scaffolding without exempting code.
    """
    if path.name != "__init__.py":
        return False
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        return False
    return len(parts) < len((ROOT_PACKAGE, *MODULE_PREFIX)) + 3


def _classify_parts(parts: tuple[str, ...]) -> tuple[str, str] | None:
    """Map path or dotted segments to ``(module, layer)``, or ``None`` when not layered code.

    Shared by both classifiers on purpose: a file and the import that targets it must agree
    on what layer they are in, and two copies of this arithmetic would eventually disagree.
    """
    head = (ROOT_PACKAGE, *MODULE_PREFIX)
    if len(parts) < len(head) + 2 or parts[: len(head)] != head:
        return None
    module, layer = parts[len(head)], parts[len(head) + 1]
    if layer not in ALLOWED_IMPORTS:
        return None
    return module, layer


def classify_file(path: Path, root: Path) -> tuple[str, str] | None:
    """Map a file to its ``(module, layer)``, or ``None`` when it is not layered code."""
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        return None
    return _classify_parts(parts)


def classify_dotted(name: str) -> tuple[str, str] | None:
    """Map a dotted import target to its ``(module, layer)``, or ``None`` when it is external."""
    return _classify_parts(tuple(name.split(".")))


def _absolute_target(node: ast.ImportFrom, path: Path, root: Path) -> str | None:
    """Resolve a possibly relative ``from ... import`` to an absolute dotted path."""
    if not node.level:
        return node.module
    try:
        package_parts = list(path.relative_to(root).parts[:-1])
    except ValueError:
        return None
    # level 1 is the containing package, level 2 its parent, and so on.
    upward = node.level - 1
    if upward:
        if upward > len(package_parts):
            return None
        package_parts = package_parts[:-upward]
    if node.module:
        package_parts.append(node.module)
    return ".".join(package_parts)


def _imported_names(tree: ast.AST, path: Path, root: Path) -> list[tuple[str, int]]:
    found: list[tuple[str, int]] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                found.append((alias.name, node.lineno))
        elif isinstance(node, ast.ImportFrom):
            target = _absolute_target(node, path, root)
            if target:
                found.append((target, node.lineno))
    return found


def _check_direction(
    origin: tuple[str, str], target: tuple[str, str], file: str, line: int
) -> Violation | None:
    _, origin_layer = origin
    _, target_layer = target
    if target_layer in ALLOWED_IMPORTS[origin_layer]:
        return None
    return Violation(
        key=f"direction:{origin_layer}->{target_layer}",
        detail=f"{origin_layer} may not import {target_layer}",
        file=file,
        line=line,
    )


def _check_slice(
    origin: tuple[str, str], target: tuple[str, str], file: str, line: int
) -> Violation | None:
    origin_module, _ = origin
    target_module, _ = target
    if target_module == origin_module or target_module in CROSS_CUTTING_MODULES:
        return None
    return Violation(
        key=f"slice:{origin_module}->{target_module}",
        detail=f"module {origin_module} reaches across into {target_module}",
        file=file,
        line=line,
    )


def _check_purity(origin: tuple[str, str], imported: str, file: str, line: int) -> Violation | None:
    _, origin_layer = origin
    if origin_layer not in PURE_LAYERS:
        return None
    distribution = imported.split(".")[0]
    if distribution not in FORBIDDEN_IN_PURE_LAYERS:
        return None
    return Violation(
        key=f"purity:{origin_layer}",
        detail=f"{origin_layer} imports framework {distribution}",
        file=file,
        line=line,
    )


def collect_violations(root: Path) -> list[Violation]:
    violations: list[Violation] = []
    for path in _iter_source_files(root):
        display = str(path.relative_to(root)).replace("\\", "/")
        origin = classify_file(path, root)
        if origin is None:
            if display in UNLAYERED_FILES or _is_package_marker(path, root):
                # Declared as outside the layer structure, or a package marker that sits above
                # the module level and therefore cannot carry a layer by construction. Both are
                # exceptions someone wrote down, which is the difference between a decision and
                # a blind spot.
                continue
            # Hard Rule 18: a file this gate cannot classify is a file this gate did not check.
            # Skipping it silently reports "clean" for code nobody looked at — which is how a
            # layout mismatch turns a whole codebase invisible while CI stays green. Either
            # teach classify_file the layout, or record the file in BASELINE deliberately.
            violations.append(
                Violation(
                    key="unclassified",
                    detail="not classifiable into (module, layer); this file was NOT checked",
                    file=display,
                    line=0,
                )
            )
            continue
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError as exc:  # a file that cannot be parsed is a failure, not a skip
            violations.append(
                Violation(
                    key="unparseable",
                    detail=f"syntax error: {exc.msg}",
                    file=display,
                    line=exc.lineno or 0,
                )
            )
            continue
        for imported, line in _imported_names(tree, path, root):
            purity = _check_purity(origin, imported, display, line)
            if purity:
                violations.append(purity)
            target = classify_dotted(imported)
            if target is None:
                continue
            direction = _check_direction(origin, target, display, line)
            if direction:
                violations.append(direction)
            slicing = _check_slice(origin, target, display, line)
            if slicing:
                violations.append(slicing)
    return violations
