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
outside it is a finding with file and line. ``dotenv_data_files`` (#635)
declares, per file and with a reason, literal ``.env`` mentions that are data
and not a read; a declaration that no longer applies is a finding. By their
form alone (#626), a ``.env`` built under a test temp dir (``tmp_path / ".env"``),
a regex-escaped ``\\.env``, a ``.env`` passed to a git query that does not emit its
content (``check-ignore``, ``log``, ``ls-files``) and the prose of an assertion
message are not findings. Shell variables that a file
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
def strip_py_strings(text: str) -> str:
    """Replace Python triple-quoted string contents with spaces (Refs #618).

    The scanner uses this so a regex match inside a docstring (e.g.
    ``os.environ["X"]`` documented in a docstring, or ``.env`` mentioned
    in prose) is not treated as a real environment read or a real
    reference to the developer's ``.env`` file. Single-line strings are
    NOT stripped because they often appear as arguments that name the
    variable being read (``os.getenv("X")``); stripping them would hide
    the actual env var name from the scanner. Line numbers are preserved
    (string contents become spaces; newlines are kept) so findings still
    cite the original line.

    Note: this intentionally does not detect reads inside f-strings that
    interpolate env vars (e.g. ``f"prefix {os.environ['X']} suffix"``).
    The current PY_READ regex does not match that form either, so the
    correction does not introduce a new false negative.
    """
    out: list[str] = []
    i = 0
    n = len(text)
    in_triple: str | None = None  # '"""' or "'''" while inside
    while i < n:
        c = text[i]
        if in_triple is not None:
            if text[i:i + 3] == in_triple:
                out.append(in_triple)
                in_triple = None
                i += 3
                continue
            out.append('\n' if c == '\n' else ' ')
            i += 1
            continue
        # Triple-quoted string start
        if c in ('"', "'") and text[i:i + 3] in ('"""', "'''"):
            delim = text[i:i + 3]
            in_triple = delim
            out.append(delim)
            i += 3
            continue
        out.append(c)
        i += 1
    return ''.join(out)
SH_READ = re.compile(r"\$\{([A-Z_][A-Z0-9_]*)(?:\[|[:=+-]|\}|\b)|\$([A-Z_][A-Z0-9_]*)\b")
PY_READ = re.compile(r"os\.environ(?:\.get)?\[\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]\]|"
                     r"os\.environ\.get\(\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]|"
                     r"os\.getenv\(\s*['\"]([A-Za-z_][A-Za-z0-9_]*)['\"]")
# #635: `.env` como FICHERO es un punto con frontera de palabra a la izquierda:
# `templates.env` (acceso a atributo) y `TEMPLATES.env` no lo son. El cargador
# se detecta por su forma (`load_dotenv`, `from dotenv`, `import dotenv`), no
# por la palabra suelta: `postgresql://dotenv` es un dato, no una lectura.
# Refs #851: la detección distingue tres casos adicionales por su forma y NO
# por la palabra `dotenv` suelta: (a) `TEMPLATES.env.get_template(...)` —
# atributo `env` de un binding local (Jinja2); (b) `"dotenv_schema"` /
# `"postgresql://dotenv"` — substring `dotenv` dentro de un identificador,
# schema name o URL, dato, no lectura; (c) `\\.env` con barra de escape —
# patrón de regex, no ruta. Los casos (a) y (b) NO contienen `.env` (el
# segundo) o lo tienen con frontera de palabra a la izquierda (el primero, vía
# `(?<!\w)`); el caso (c) está cubierto por la segunda lookbehind
# `(?<!(?<!\\)\\)\.env\b`. La suite `TestDotenvPrecision` lock-downea esto
# desde #851 para que un refactor del regex no pueda regresionar.
# #626: un `\.env` con barra de escape es un patrón de regex, no una ruta.
DOTENV_FILE = re.compile(r"(?<!\w)(?<!(?<!\\)\\)\.env\b")
# #626 B: un `.env` colgado de un directorio temporal del test (`tmp_path / ".env"`)
# es un fixture propio, no el `.env` del desarrollador.
TMP_DOTENV = re.compile(r"\b(?:Path\(\s*)?tmp\w*\s*\)?\s*/\s*['\"]\.env['\"]")
# #626 C: consultar a git por el `.env` (ignorado, historial, índice) no lee su
# contenido. `git show HEAD:.env` sí lo emite y sigue siendo hallazgo.
GIT_META = re.compile(r"(?<![^\W_])check[_-]ignored?\b|\bgit\b['\",\s]+(?:-C['\",\s]+\S+['\",\s]+)?"
                      r"(?:log|ls-files|status|add|rm)\b")
# #626 C: el mensaje en prosa de una aserción o de un fallo (literal con espacios).
ASSERT_LINE = re.compile(r"^\s*(?:assert|raise|pytest\.fail|self\.(?:fail|assert\w+))\b")
PROSE_LITERAL = re.compile(r"(['\"])(?=[^'\"]*\s)[^'\"]*\1")


def dotenv_file_ref(line: str, python: bool) -> bool:
    r"""True si la línea nombra el ``.env`` local como fichero (#626, #851).

    Es un detector POR FORMA: un ``.env`` con frontera de palabra a la
    izquierda (``(?<!\w)``) y sin barra de escape precedente. Tres
    formulaciones que la regex maneja correctamente y que cualquier
    refactor del regex debe preservar (Refs #851):

    - ``TEMPLATES.env.get_template(template_name)`` -- atributo ``env`` de
      un binding local (Jinja2, dataclass, etc.). El carácter precedente
      es una letra, así que la lookbehind ``(?<!\w)`` falla y la línea no
      se marca.
    - ``"dotenv_schema"``, ``"postgresql://dotenv"`` -- ``dotenv`` como
      substring de un identificador o URL. No hay ``.env`` literal, así
      que DOTENV_FILE no matchea; el cargador se detecta por
      DOTENV_LOADER, no por la palabra suelta.
    - ``\.env`` en código — patrón de regex escapado, no ruta. La
      segunda lookbehind (``(?<!(?<!\)\)``) exige que NO haya una única
      barra de escape precediendo al ``.env``.
    """

    if GIT_META.search(line):
        return False
    if python:
        line = TMP_DOTENV.sub(" ", line)
        if ASSERT_LINE.search(line):
            line = PROSE_LITERAL.sub(" ", line)
    return bool(DOTENV_FILE.search(line))
DOTENV_LOADER = re.compile(r"\bload_dotenv\b|(?:^|\s)(?:from|import)\s+dotenv\b")
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


def load_allowlist(path: str) -> tuple:
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
    return set(doc["env_vars"]), set(excluded), load_dotenv_data(doc)


def load_dotenv_data(doc: dict) -> dict:
    """#635: `dotenv_data_files` declara, fichero a fichero y con motivo, las
    referencias literales a `.env` que son DATO (fixtures que escriben un
    `.env` en tmp_path, aserciones sobre `.gitignore`), no lectura. Solo eximen
    la mención de `.env`: un cargador o una lectura de entorno sigue siendo
    hallazgo en el mismo fichero."""
    entries = doc.get("dotenv_data_files", [])
    if not isinstance(entries, list) or not all(
            isinstance(e, dict) and isinstance(e.get("file"), str) and e["file"]
            and isinstance(e.get("reason"), str) and e["reason"].strip()
            for e in entries):
        raise GateError('"dotenv_data_files" must be a list of {"file", "reason"} '
                        'objects: a relative file path and a non-empty reason')
    files = {e["file"] for e in entries}
    if len(files) != len(entries):
        raise GateError('"dotenv_data_files" repeats a file')
    return {e["file"]: e["reason"] for e in entries}


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


def scan_sh(path: Path, rel: str, allow: set, report: list, data_ok: bool) -> bool:
    text = path.read_text(encoding="utf-8", errors="replace")
    # Locals are file-wide: a variable assigned anywhere in the file is a
    # shell local, never an environment read (HR-1: structured, provable).
    assigned = set(ASSIGNED.findall(text)) | set(FOR_VAR.findall(text)) \
        | set(ASSIGN_CMD.findall(text)) \
        | set(re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)=\(", text))
    found_ref = False
    for number, raw in enumerate(text.splitlines(), 1):
        # Comments are stripped (quote-aware) before scanning: a ${VAR:-x}
        # mentioned in a comment is documentation, not an environment read.
        line = strip_comment(raw)
        file_ref = dotenv_file_ref(line, False) and not DOTENV_EXEMPT.search(line)
        found_ref = found_ref or file_ref
        if DOTENV_LOADER.search(line) or (file_ref and not data_ok):
            report.append(f"HR-28 {rel}:{number}: references a local .env file — "
                          "the suite must not read the developer's local environment")
        for m in SH_READ.finditer(line):
            name = m.group(1) or m.group(2)
            if name in assigned or name in allow:
                continue
            report.append(f"HR-28 {rel}:{number}: reads environment variable "
                          f"{name!r} not declared in the versioned allowlist")
    return found_ref


def scan_py(path: Path, rel: str, allow: set, report: list, data_ok: bool) -> bool:
    # Refs #618: prose inside Python string literals (docstrings, regular
    # strings) is not a real env read or a real reference to the
    # developer's ``.env``. Strip string contents before applying the
    # regex surface. ``strip_comment`` is still applied per-line
    # afterwards to drop ``#`` comments.
    text = path.read_text(encoding="utf-8", errors="replace")
    stripped = strip_py_strings(text)
    found_ref = False
    for number, raw in enumerate(stripped.splitlines(), 1):
        # La exención solo ampara la coincidencia .env de plantilla; las
        # lecturas de os.environ de la misma línea se siguen escaneando.
        # El strip de comentarios es quote-aware también aquí.
        line = strip_comment(raw)
        file_ref = dotenv_file_ref(line, True) and not DOTENV_EXEMPT.search(line)
        found_ref = found_ref or file_ref
        if DOTENV_LOADER.search(line) or (file_ref and not data_ok):
            report.append(f"HR-28 {rel}:{number}: references a local .env file — "
                          "the suite must not read the developer's local environment")
        for m in PY_READ.finditer(line):
            name = m.group(1) or m.group(2) or m.group(3)
            if name in allow:
                continue
            # Refs #618: ``os.environ["X"] = value`` (LHS of plain ``=``)
            # is a write, not a read. Skip it. Augmented assignments
            # (``+=``, ``-=``, ...) and equality (``==``) keep being
            # treated as reads.
            after = line[m.end():].lstrip()
            if after and after[0] == '=' and (len(after) < 2 or after[1] != '='):
                continue
            report.append(f"HR-28 {rel}:{number}: reads environment variable "
                          f"{name!r} via os.environ/getenv not declared in the "
                          "versioned allowlist")
    return found_ref


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv):
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Env isolation gate (HR-28).")
    ap.add_argument("--root", required=True)
    ap.add_argument("--allowlist", required=True)
    args = ap.parse_args(argv)
    try:
        root = Path(args.root)
        if not root.is_dir():
            raise GateError(f"root {str(root)!r} is not a directory (empty subject)")
        allow, excluded, dotenv_data = load_allowlist(args.allowlist)
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
    refs = set()
    for path in files:
        rel = path.relative_to(root).as_posix()
        scan = scan_py if path.suffix == ".py" else scan_sh
        if scan(path, rel, allow, report, rel in dotenv_data):
            refs.add(rel)
    # #635: una exención que ya no hace falta falla, para que la lista solo encoja.
    for rel in sorted(set(dotenv_data) - refs):
        report.append(f"HR-28 STALE: dotenv_data_files declares {rel!r} but it no longer "
                      "references .env (or is not a suite file): remove the entry")
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
