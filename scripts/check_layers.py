#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — scripts/check_layers.py
"""Hexagonal layer gate: dependency direction, vertical slicing, and layer purity.

Stdlib only, so the gate runs before the project has installed anything. Walks the AST of every
module under the root package, resolves each import to a ``(module, layer)`` pair, and fails on
any edge the architecture does not allow.

Regex cannot do this job: it breaks on line continuations, comments, aliased imports, and relative
imports. Hard Rule 9 requires the AST.

The root package for the Lanzadera MVP is ``app.src.modules``. Each module lives as a direct
sub-package (``app.src.modules.<module>``) and each layer as a sub-sub-package (``domain``,
``ports``, ``application``, ``adapters``, ``shared``, ``di``, ``delivery``). The script walks
``<root>/app/src/modules/<module>/<layer>/...`` and classifies every ``.py`` file.

Exit codes:
    0  no violation outside BASELINE
    1  a new violation, a BASELINE entry above its recorded count, or a BASELINE entry past its
       target date while still above target
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION — adjust this block when instantiating. Everything below it is mechanism.
# --------------------------------------------------------------------------------------------

# Dotted path to the package root. The gate walks `<root>/app/src/modules/...`.
ROOT_PACKAGE = "app.src.modules"
ROOT_PACKAGE_PARTS = ROOT_PACKAGE.split(".")

#: Modules every other module may depend on. Keep this set as small as it can possibly be; each
#: entry is a hole in the vertical slicing rule.
CROSS_CUTTING_MODULES = frozenset({"shared"})

#: Which layers a given layer may import from. A layer absent from this table is not a layer.
ALLOWED_IMPORTS: dict[str, frozenset[str]] = {
    "domain": frozenset({"domain", "shared"}),
    "ports": frozenset({"domain", "ports", "shared"}),
    "application": frozenset({"domain", "ports", "application", "shared"}),
    "adapters": frozenset({"domain", "ports", "adapters", "shared"}),
    "shared": frozenset({"domain", "ports", "shared"}),
    "delivery": frozenset(
        {"domain", "ports", "application", "adapters", "shared", "delivery"}
    ),
    "di": frozenset(
        {"domain", "ports", "application", "adapters", "shared", "delivery", "di"}
    ),
}

#: Layers that must not touch a framework at all.
PURE_LAYERS = frozenset({"domain", "ports", "application"})

#: Top-level distributions banned inside PURE_LAYERS.
FORBIDDEN_IN_PURE_LAYERS = frozenset(
    {"fastapi", "sqlalchemy", "alembic", "jinja2", "httpx", "asyncpg", "starlette", "pydantic_settings"}
)

#: Paths excluded from the walk. Hard Rule 14: this is the declared untestable/ungoverned
#: boundary. Keep it thin and keep it honest — every entry here is code nobody is checking.
EXCLUDED_PARTS = frozenset({"__pycache__", ".venv", "venv", "build", "dist", "migrations"})


@dataclass(frozen=True)
class BaselineEntry:
    """A tolerated violation count with a mandatory exit plan.

    Hard Rule 12: a ratchet is a ramp toward a ceiling, never the destination. ``target`` is the
    value this key must reach; ``target_date`` is when the tolerance expires. After that date the
    gate fails even if the count never grew, which is what stops a ratchet from becoming permanent.
    """

    count: int
    target: int
    target_date: str  # ISO-8601, YYYY-MM-DD


#: Shrink-only. Lowering a count is a lock-in and is allowed. Raising one requires editing this
#: file in the same PR that raises it, which is the point.
BASELINE: dict[str, BaselineEntry] = {}

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class Violation:
    key: str
    detail: str
    file: str
    line: int


def _package_dir(root: Path) -> Path:
    """Resolve ``<root>/<ROOT_PACKAGE>`` where ``ROOT_PACKAGE`` is a dotted path."""
    return root.joinpath(*ROOT_PACKAGE_PARTS)


#: Composition-root files outside the module/layer lattice. They sit above
#: the hexagonal inversion; the gate intentionally does not classify them.
#: Phase 0 ships exactly one such file (`app/src/main.py`); Phase 4 will hoist
#: the FastAPI composition root into `app/src/modules/lanzadera/delivery/...`
#: and remove this entry.
COMPOSITION_ROOTS: tuple[str, ...] = ("app/src/main.py",)


def _iter_source_files(root: Path) -> list[Path]:
    package_root = _package_dir(root)
    if not package_root.is_dir():
        return []
    files: list[Path] = []
    for path in sorted(package_root.rglob("*.py")):
        if EXCLUDED_PARTS.intersection(path.parts):
            continue
        files.append(path)
    return files


def _is_composition_root(path: Path, root: Path) -> bool:
    """Composition-root files sit above the hexagonal inversion."""
    try:
        rel = path.relative_to(root).as_posix()
    except ValueError:
        return False
    return rel in COMPOSITION_ROOTS


# Sentinel returned by `classify_file` for files that are intentionally
# outside the layered lattice (composition roots, empty `__init__.py`
# markers, etc.). `collect_violations` filters these before recording any
# "unclassified" violation.
SKIP_SENTINEL = ("__skip__", "")


def _should_skip(path: Path, root: Path) -> bool:
    """True for files intentionally outside the module/layer lattice."""
    if _is_composition_root(path, root):
        return True
    # Empty `__init__.py` files at any level are package markers, not code.
    # They sit on the path but contribute zero statements to coverage, so the
    # gate has nothing to enforce.
    if path.name == "__init__.py":
        return True
    return False


def classify_file(path: Path, root: Path) -> tuple[str, str] | None:
    """Map a file to its ``(module, layer)``, or ``None`` when it is not layered code.

    Returns ``SKIP_SENTINEL`` for files intentionally outside the lattice
    (composition roots, empty `__init__.py` markers). The collector filters
    those before reporting violations.
    """
    if _should_skip(path, root):
        return SKIP_SENTINEL
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        return None
    # `<root>/app/src/modules/<module>/<layer>/<file>.py` — at least 6 parts
    # (`app`, `src`, `modules`, `<module>`, `<layer>`, `<file>.py`).
    if len(parts) < 6:
        return None
    if list(parts[: len(ROOT_PACKAGE_PARTS)]) != ROOT_PACKAGE_PARTS:
        return None
    module = parts[len(ROOT_PACKAGE_PARTS)]
    layer = parts[len(ROOT_PACKAGE_PARTS) + 1]
    if layer not in ALLOWED_IMPORTS:
        return None
    return module, layer


def classify_dotted(name: str) -> tuple[str, str] | None:
    """Map a dotted import target to its ``(module, layer)``, or ``None`` when it is external."""
    parts = name.split(".")
    if len(parts) <= len(ROOT_PACKAGE_PARTS):
        return None
    if parts[: len(ROOT_PACKAGE_PARTS)] != ROOT_PACKAGE_PARTS:
        return None
    # A target like `app.src.modules.lanzadera.domain.user`:
    # module = lanzadera, layer = domain.
    # A target like `app.src.shared.cache` has no layer — it is a cross-cutting utility,
    # classified as belonging to `shared` so the direction rules still apply.
    if len(parts) == len(ROOT_PACKAGE_PARTS) + 1:
        # Exactly the module or `shared` — treat as cross-cutting.
        return parts[len(ROOT_PACKAGE_PARTS)], "shared"
    module = parts[len(ROOT_PACKAGE_PARTS)]
    layer = parts[len(ROOT_PACKAGE_PARTS) + 1]
    if layer not in ALLOWED_IMPORTS:
        # Allow `app.src.shared.<x>` as a `shared` import target.
        if module == "shared":
            return module, "shared"
        return None
    return module, layer


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


def _check_purity(
    origin: tuple[str, str], imported: str, file: str, line: int
) -> Violation | None:
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


def collect_violations(root: Path) -> tuple[list[Violation], int]:
    """Return ``(violations, skipped_count)``.

    ``skipped_count`` is the number of files that were intentionally outside
    the layered lattice (composition roots, empty `__init__.py` markers).
    Hard Rule 18 still holds for files the gate cannot classify for an
    unintentional reason — those are violations, not skips.
    """
    violations: list[Violation] = []
    skipped = 0
    for path in _iter_source_files(root):
        display = str(path.relative_to(root)).replace("\\", "/")
        origin = classify_file(path, root)
        if origin == SKIP_SENTINEL:
            # Composition root or empty `__init__.py` — intentionally outside
            # the module lattice. Silent skip, not a violation.
            skipped += 1
            continue
        if origin is None:
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
    return violations, skipped


def evaluate(violations: list[Violation], today: date) -> tuple[int, list[str]]:
    """Compare observed violations against BASELINE. Returns ``(exit_code, report_lines)``."""
    observed: dict[str, list[Violation]] = {}
    for violation in violations:
        observed.setdefault(violation.key, []).append(violation)

    lines: list[str] = []
    failed = False

    for key in sorted(observed):
        entries = observed[key]
        allowance = BASELINE.get(key)
        if allowance is None:
            failed = True
            lines.append(f"FAIL  {key}: {len(entries)} occurrence(s), not in BASELINE")
        elif len(entries) > allowance.count:
            failed = True
            lines.append(
                f"FAIL  {key}: {len(entries)} occurrence(s), BASELINE allows {allowance.count}"
            )
        elif today.isoformat() > allowance.target_date and len(entries) > allowance.target:
            failed = True
            lines.append(
                f"FAIL  {key}: BASELINE expired on {allowance.target_date} with "
                f"{len(entries)} occurrence(s), target was {allowance.target}"
            )
        elif len(entries) < allowance.count:
            lines.append(
                f"NOTE  {key}: {len(entries)} occurrence(s), below BASELINE {allowance.count}; "
                f"lower the BASELINE to lock the gain in"
            )
        for entry in entries:
            lines.append(f"        {entry.file}:{entry.line}  {entry.detail}")

    for key in sorted(BASELINE):
        if key in observed:
            continue
        lines.append(f"NOTE  {key}: no longer occurs; remove it from BASELINE")

    if not failed:
        lines.append("OK    layer gate clean")
    return (1 if failed else 0), lines


def build_report(
    violations: list[Violation], status: str, files_seen: int = 0, skipped: int = 0
) -> dict:
    unclassified = sum(1 for violation in violations if violation.key == "unclassified")
    return {
        "gate": "layers",
        "status": status,
        "indicators": {
            "violations": len(violations),
            "violation_classes": len({violation.key for violation in violations}),
            # Coverage of the gate itself: how much of the package it actually inspected.
            "files_checked": files_seen - unclassified - skipped,
            "files_unclassified": unclassified,
            "files_skipped": skipped,
        },
        "ceilings": {"violations": 0, "violation_classes": 0, "files_unclassified": 0},
        "findings": [
            {"file": violation.file, "line": violation.line, "detail": violation.detail}
            for violation in sorted(violations, key=lambda item: (item.file, item.line, item.key))
        ],
    }


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python picks the output encoding from the platform locale, so the same gate emits different
    bytes on a Windows workstation (cp1252) and a Linux runner (utf-8) — and a non-encodable
    character crashes the write outright. A harness that claims determinism cannot let its own
    output depend on where it ran.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="repository root containing the package directory (default: cwd)",
    )
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)

    root = args.root.resolve()
    package_dir = _package_dir(root)
    if not package_dir.is_dir():
        message = f"root package '{ROOT_PACKAGE}' not found under {root} (looked in '{package_dir}')"
        if args.json:
            print(json.dumps({"gate": "layers", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    violations, skipped = collect_violations(root)
    files_seen = len(_iter_source_files(root))
    exit_code, lines = evaluate(violations, date.today())

    if args.json:
        status = "pass" if exit_code == 0 else "fail"
        print(json.dumps(build_report(violations, status, files_seen, skipped), indent=2))
    else:
        for line in lines:
            print(line)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())