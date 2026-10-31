#!/usr/bin/env python3
# ci-pattern asset — clasificador de alcance por suite (#1440)
"""Decide qué suites puede romper una PR a partir de su lista de ficheros.

Cada suite declara en ``.github/change-scope.json`` (el de la BASE de la PR) la
lista **positiva** de rutas que pueden romperla, como hace
``Gentleman-Programming/gentle-ai`` con ``on.pull_request.paths``. El script
publica ``<suite>_needed=true|false`` por suite; el agregador ``required``
acepta el salto de una suite solo a través de un ``marker_lanes`` con
``equals: "false"`` (``assets/required-jobs``).

Falla cerrado (HR-3): sale con 2 y sin salida parcial si la declaración o la
lista de ficheros no se pueden leer; marca todas las suites necesarias si el
diff está vacío, si una ruta llega entrecomillada por ``git`` (no se puede
casar) o si toca una ruta de :data:`IMPLICIT_ALWAYS`, de ``always`` o un
fichero sellado del manifiesto de gobierno.

El módulo no ejecuta ``git`` ni lanza procesos: el workflow le pasa la salida
de ``git diff --no-renames --name-only BASE...HEAD`` en un fichero. Los
patrones siguen la semántica de ``.gitattributes`` de ``vendored_budget``.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# Única regla de patrones del canon (#931): la misma que el presupuesto.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "budget"))
from vendored_budget import _gitattributes_regex  # noqa: E402

SCOPE_VERSION = 1
SCOPE_KEYS = frozenset({"version", "suites", "always"})
SUITE_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]*\Z")

#: Rutas que siempre exigen todas las suites: el canon vendorizado, la propia
#: declaración y el manifiesto de gobierno. Ninguna declaración las saca.
IMPLICIT_ALWAYS = (".agents/skills/ci-pattern/assets/**", ".github/change-scope.json",
                   ".governance-manifest.json")


class UsageError(Exception):
    """Entrada ilegible: el script sale con 2 sin publicar nada."""


class InvalidScopeError(UsageError):
    """La declaración no cumple el formato; ``code`` elige el mensaje."""

    MESSAGES = {
        "patterns": "{0} must be a non-empty list of path patterns",
        "pattern": "{0} holds a non-string or empty pattern",
        "unsupported": "{0}: {1!r} is not supported (no negation; use 'dir/**')",
        "json": "not valid JSON ({0})",
        "keys": "must be an object with only {0}",
        "version": "version must be {0}",
        "suites": "suites must be a non-empty object",
        "name": "suite name {0!r} must match {1}",
    }

    def __init__(self, code: str, *values: object) -> None:
        super().__init__("invalid scope: " + self.MESSAGES[code].format(*values))


class UnreadableInputError(UsageError):
    def __init__(self, flag: str, path: str, exc: Exception) -> None:
        super().__init__(f"cannot read {flag} {path}: {exc}")


def _validate_patterns(patterns: object, where: str) -> list[str]:
    """Una lista no vacía de patrones de ruta, sin negación ni barra final."""
    if not isinstance(patterns, list) or not patterns:
        raise InvalidScopeError("patterns", where)
    for pattern in patterns:
        if not isinstance(pattern, str) or not pattern.strip():
            raise InvalidScopeError("pattern", where)
        if pattern.startswith("!") or pattern.endswith("/"):
            raise InvalidScopeError("unsupported", where, pattern)
    return patterns


def load_scope(text: str) -> tuple[dict[str, list[str]], list[str]]:
    """Valida la declaración y devuelve ``(suites, always)``."""
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise InvalidScopeError("json", exc) from exc
    if not isinstance(data, dict) or set(data) - SCOPE_KEYS:
        raise InvalidScopeError("keys", sorted(SCOPE_KEYS))
    if data.get("version") != SCOPE_VERSION:
        raise InvalidScopeError("version", SCOPE_VERSION)
    suites = data.get("suites")
    if not isinstance(suites, dict) or not suites:
        raise InvalidScopeError("suites")
    for name, patterns in suites.items():
        if not SUITE_NAME.match(name):
            raise InvalidScopeError("name", name, SUITE_NAME.pattern)
        _validate_patterns(patterns, f"suites.{name}")
    always = _validate_patterns(data["always"], "always") if "always" in data else []
    return suites, always


def _matches(path: str, patterns: list[str]) -> bool:
    return any(_gitattributes_regex(p).match(path) for p in patterns)


def classify(changed: list[str], suites: dict[str, list[str]], always: list[str],
             sealed: set[str]) -> dict[str, bool]:
    """``{suite: needed}``. Cualquier señal de duda marca todas las suites."""
    every = [*IMPLICIT_ALWAYS, *always]
    if not changed or any(
            p.startswith('"') or p in sealed or _matches(p, every) for p in changed):
        return dict.fromkeys(suites, True)
    return {name: any(_matches(p, patterns) for p in changed)
            for name, patterns in suites.items()}


def _read_lines(flag: str, path: str) -> list[str]:
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise UnreadableInputError(flag, path, exc) from exc
    return [line.strip() for line in text.splitlines() if line.strip()]


def _read_scope(path: str) -> tuple[dict[str, list[str]], list[str]]:
    try:
        text = Path(path).read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise UnreadableInputError("--scope", path, exc) from exc
    return load_scope(text)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--changed-files", required=True,
                        help="output of git diff --no-renames --name-only BASE...HEAD")
    parser.add_argument("--scope", required=True, help=".github/change-scope.json of the BASE")
    parser.add_argument("--sealed-files", help="sealed paths of the BASE manifest, one per line")
    parser.add_argument("--github-output", help="file to append the outputs to ($GITHUB_OUTPUT)")
    args = parser.parse_args(argv)
    try:
        suites, always = _read_scope(args.scope)
        changed = _read_lines("--changed-files", args.changed_files)
        sealed = set(_read_lines("--sealed-files", args.sealed_files)) if args.sealed_files else set()
    except UsageError as exc:
        print(f"classify_changes: {exc}", file=sys.stderr)
        return 2
    lines = [f"{name}_needed={'true' if needed else 'false'}"
             for name, needed in classify(changed, suites, always, sealed).items()]
    report = "".join(f"{line}\n" for line in lines)
    if args.github_output:
        with Path(args.github_output).open("a", encoding="utf-8") as handle:
            handle.write(report)
    sys.stdout.write(report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
