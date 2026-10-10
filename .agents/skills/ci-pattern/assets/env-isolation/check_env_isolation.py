#!/usr/bin/env python3
# ci-pattern asset — env isolation gate (HR-28, DysTelefonica/team-skills#245)
"""Suite environment isolation gate (HR-28).

Scans a suite tree (``--root``, recursively, ``*.sh`` and ``*.py``) for two
kinds of leaks from the developer's local environment:

- a reference to a local ``.env`` file (``source .env``, ``load_dotenv()``,
  ``cat .env`` ...). Template names (``.env.example``, ``.env.sample``,
  ``.env.template``) are not local state and are exempt;
- an environment variable read that is neither assigned locally in the same
  file (shell assignments, ``for`` loop variables, ``local``/``readonly``/
  ``declare``/``export``) nor declared in the versioned allowlist.

The allowlist (``--allowlist``, versioned JSON: ``{"env_vars": [...]}``) is
the contract of what the suites may read from the environment; anything
outside it is a finding with file and line. Shell variables that a file
assigns are locals, not environment reads; a shell variable that is read
but never assigned in its file is treated as an environment read (fail
closed: the scanner cannot prove it is local, HR-1).

Exit codes: 0 clean · 1 findings · 2 invalid input (missing root, missing
or malformed allowlist, empty subject — a tree with no suite files).
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ASSIGNED = re.compile(r"(?:^|[\s;&(])([A-Za-z_][A-Za-z0-9_]*)=")
ASSIGN_CMD = re.compile(r"\b(?:mapfile|read|local|readonly|declare|export)"
                        r"(?:\s+-\S+)*\s+([A-Za-z_][A-Za-z0-9_]*)")
FOR_VAR = re.compile(r"\bfor\s+([A-Za-z_][A-Za-z0-9_]*)\b")
def strip_comment(line: str) -> str:
    """Quote-aware comment strip (R2, auditor de #259): un # dentro de
    comillas simples o dobles es contenido, no comentario."""
    quote = None
    for i, c in enumerate(line):
        if quote:
            quote = None if c == quote else quote
        elif c in "'\"":
            quote = c
        elif c == "#" and (i == 0 or line[i - 1] in " \t"):
            return line[:i]
    return line
SH_READ = re.compile(r"\$\{([A-Z_][A-Z0-9_]*)(?:\[|[:=+-]|\}|\b)|\$([A-Z_][A-Z0-9_]*)\b")
PY_READ = re.compile(r"os\.environ(?:\.get)?\[\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]|"
                     r"os\.environ\.get\(\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]|"
                     r"os\.getenv\(\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]")
DOTENV = re.compile(r"\.env\b|\bdotenv\b|load_dotenv")
DOTENV_EXEMPT = re.compile(r"\.env\.(?:example|sample|template)\b")


# #315 (HR-28): el DOMINIO del gate —qué extensiones escanea— es declarado, no
# implícito. Un árbol con ficheros de suite de otro lenguaje no se ignora en
# silencio: se nombra, y solo pasa si el contrato lo excluye explícitamente
# (`excluded_extensions` del allowlist) o si el dominio se amplía con su
# escáner. Un sujeto incompleto no es un pase (HR-3).
DOMAIN_EXTENSIONS = (".sh", ".py")
# Extensiones que son CÓDIGO y por tanto exigen una decisión explícita. Un
# `.md`, un `.json` de datos o un `.env.example` no son sujeto de este gate.
SOURCE_EXTENSIONS = (".sh", ".py", ".php", ".ts", ".tsx", ".js", ".mjs", ".cjs",
                     ".rb", ".go", ".java", ".cs", ".kt", ".pl", ".ps1", ".lua",
                     ".r", ".swift", ".scala", ".ex", ".exs", ".dart", ".rs",
                     ".vb", ".cls", ".bas")


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def load_allowlist(path: str) -> set:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GateError(f"allowlist {path!r} unreadable: {exc}") from exc
    if not isinstance(doc, dict) or not isinstance(doc.get("env_vars"), list) \
            or not all(isinstance(v, str) and v for v in doc["env_vars"]):
        raise GateError('allowlist must be an object with an "env_vars" list of names')
    excluded = doc.get("excluded_extensions", [])
    if not isinstance(excluded, list) or not all(
            isinstance(x, str) and x.startswith(".") for x in excluded):
        raise GateError('"excluded_extensions" must be a list of extensions '
                        'starting with "." (explicit exclusion of a language)')
    return set(doc["env_vars"]), set(excluded)


def suite_files(root: Path) -> list[Path]:
    files = [p for p in sorted(root.rglob("*"))
             if p.is_file() and p.suffix in DOMAIN_EXTENSIONS]
    if not files:
        raise GateError(
            f"no {' or '.join('*' + e for e in DOMAIN_EXTENSIONS)} suite files under "
            f"{str(root)!r} (empty subject)")
    return files


def out_of_domain_files(root: Path, excluded: set) -> list[Path]:
    """Ficheros de CÓDIGO fuera del dominio declarado y no excluidos."""
    return [p for p in sorted(root.rglob("*"))
            if p.is_file() and p.suffix in SOURCE_EXTENSIONS
            and p.suffix not in DOMAIN_EXTENSIONS and p.suffix not in excluded]


def domain_finding(root: Path, fuera: list[Path]) -> str:
    por_extension = {}
    for path in fuera:
        por_extension.setdefault(path.suffix, []).append(
            path.relative_to(root).as_posix())
    detalle = ", ".join(f"{ext} ({len(paths)})"
                        for ext, paths in sorted(por_extension.items()))
    ejemplos = ", ".join(path.relative_to(root).as_posix() for path in fuera[:3])
    return (f"HR-28 DOMAIN: {len(fuera)} suite file(s) in a language outside the gate's "
            f"declared domain: {detalle} — e.g. {ejemplos}. This gate scans "
            f"{', '.join(DOMAIN_EXTENSIONS)}: declare the exclusion explicitly in the "
            "contract (`excluded_extensions` of the allowlist) or extend the domain with "
            "its own scanner; an incomplete subject is not a pass (HR-3).")


def scan_sh(path: Path, rel: str, allow: set, report: list) -> None:
    text = path.read_text(encoding="utf-8", errors="replace")
    # Locals are file-wide: a variable assigned anywhere in the file is a
    # shell local, never an environment read (HR-1: structured, provable).
    assigned = set(ASSIGNED.findall(text)) | set(FOR_VAR.findall(text)) \
        | set(ASSIGN_CMD.findall(text)) \
        | set(re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)=\(", text))
    for number, raw in enumerate(text.splitlines(), 1):
        # Comments are stripped (quote-aware) before scanning: a ${VAR:-x}
        # mentioned in a comment is documentation, not an environment read.
        line = strip_comment(raw)
        if DOTENV.search(line) and not DOTENV_EXEMPT.search(line):
            report.append(f"HR-28 {rel}:{number}: references a local .env file — "
                          "the suite must not read the developer's local environment")
        for m in SH_READ.finditer(line):
            name = m.group(1) or m.group(2)
            if name in assigned or name in allow:
                continue
            report.append(f"HR-28 {rel}:{number}: reads environment variable "
                          f"{name!r} not declared in the versioned allowlist")


def scan_py(path: Path, rel: str, allow: set, report: list) -> None:
    for number, raw in enumerate(path.read_text(encoding="utf-8",
                                                 errors="replace").splitlines(), 1):
        # La exención solo ampara la coincidencia .env de plantilla; las
        # lecturas de os.environ de la misma línea se siguen escaneando.
        # El strip de comentarios es quote-aware también aquí.
        line = strip_comment(raw)
        if DOTENV.search(line) and not DOTENV_EXEMPT.search(line):
            report.append(f"HR-28 {rel}:{number}: references a local .env file — "
                          "the suite must not read the developer's local environment")
        for m in PY_READ.finditer(line):
            name = m.group(1) or m.group(2) or m.group(3)
            if name in allow:
                continue
            report.append(f"HR-28 {rel}:{number}: reads environment variable "
                          f"{name!r} via os.environ/getenv not declared in the "
                          "versioned allowlist")


def main(argv):
    ap = argparse.ArgumentParser(description="Env isolation gate (HR-28).")
    ap.add_argument("--root", required=True)
    ap.add_argument("--allowlist", required=True)
    args = ap.parse_args(argv)
    try:
        root = Path(args.root)
        if not root.is_dir():
            raise GateError(f"root {str(root)!r} is not a directory (empty subject)")
        allow, excluded = load_allowlist(args.allowlist)
        fuera = out_of_domain_files(root, excluded)
        try:
            files = suite_files(root)
        except GateError:
            # Un árbol con SOLO ficheros fuera del dominio no es duda: es el
            # hallazgo de dominio, y se nombra.
            if not fuera:
                raise
            files = []
    except GateError as exc:
        print(f"ENV-ISOLATION ERROR: {exc}", file=sys.stderr)
        return 2
    report = [domain_finding(root, fuera)] if fuera else []
    for path in files:
        rel = path.relative_to(root).as_posix()
        (scan_py if path.suffix == ".py" else scan_sh)(path, rel, allow, report)
    if report:
        print(f"HR-28 FAIL: {len(report)} finding(s) — suite environment is not "
              "isolated from the local environment")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"ENV-ISOLATION OK: {len(files)} suite file(s) scanned — no .env reads, "
          "no undeclared environment variables.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
