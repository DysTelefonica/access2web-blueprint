#!/usr/bin/env python3
# ci-pattern asset — fuente única del presupuesto de revisión (#931)
"""Regla de presupuesto compartida por ``pr-size`` y ``pr-contract`` (A1).

Las rutas de :data:`VENDORED_ROOTS` quedan fuera del presupuesto de revisión,
pero un fichero renombrado desde una ruta de autoría hacia una vendorizada
sigue contando, aunque el renombrado ocurra en otro commit del mismo PR.

Desde #1088 la lista tiene una sola entrada: ``.agents/skills/``, el árbol
vendorizado que los consumers reciben del canon. ``skills/`` **no** se
excluye: en el canon es el catálogo de autoría (64 skills escritas aquí) y en
un consumer es su catálogo propio, así que sus líneas son trabajo que alguien
tiene que revisar. Excluirlo hacía que el gate midiera ``changed_lines: 0``
en PRs que reescribían una skill entera de 569 líneas.

El módulo no ejecuta ``git`` ni lanza procesos: el activo se copia
byte a byte en consumers con ratchet de ruff extendido (S603/S607) y quien
llama ya tiene su runner, así que lo inyecta (#753). El contrato es
``run_git(*args) -> str``, ligado a su repositorio, que devuelve stdout y
lanza ``RuntimeError`` si el comando falla; un ``git`` fallido deja el origen
desconocido, que es lo que cuenta hacia dentro del presupuesto.

``origins_for`` arma el fichero de origen del PR que ``pr-contract`` recibe
con ``--origins-file`` (no ejecuta ``git``, HR-3): una línea por fichero
vendorizado del PR, ``<ruta-nueva>TAB<origen>``, donde ``<origen>`` es la ruta
de autoría, o ``-`` si el origen es vendorizado o desconocido.
"""

from __future__ import annotations

import re
from collections.abc import Callable

#: Runner de git inyectado: recibe los argumentos, devuelve stdout, lanza ``RuntimeError`` si falla.
RunGit = Callable[..., str]

#: Única definición de las rutas excluidas del presupuesto de revisión.
#:
#: Solo ``.agents/skills/``: el árbol vendorizado de los consumers, cuyo
#: código no escribió quien abre la PR. ``skills/`` no está aquí (#1088):
#: en el canon es el catálogo de autoría y en un consumer es su catálogo
#: propio, y en los dos casos sus líneas cuentan para el presupuesto.
VENDORED_ROOTS = (".agents/skills/",)

#: Campos de una línea de ``--name-status``: estado y ruta; renombrado: estado, origen y destino.
_ADDED_FIELDS = 2
_RENAME_FIELDS = 3

#: Forma ``prefijo{old => new}sufijo`` de ``git diff --numstat``.
_RENAME_BRACES_RE = re.compile(r"\{(.*?) => (.*?)\}")


def is_under_vendored(path: str) -> bool:
    """True cuando ``path`` cae bajo una de :data:`VENDORED_ROOTS`."""
    return any(path.startswith(root) for root in VENDORED_ROOTS)


def counts_to_budget(old_path: str, new_path: str) -> bool:
    """Decide si los cambios de un fichero cuentan para el presupuesto.

    1. Destino fuera de las rutas vendorizadas: cuenta.
    2. Destino vendorizado y sin renombrado desde autoría: no cuenta.
    3. Renombrado desde una ruta de autoría hacia una vendorizada: cuenta.
    """
    if not is_under_vendored(new_path):
        return True
    return old_path != new_path and not is_under_vendored(old_path)


def _strip_diff_prefix(path: str) -> str:
    """Quita el prefijo ``a/`` o ``b/`` que git usa en la salida de diff."""
    if path.startswith("a/") or path.startswith("b/"):
        return path[2:]
    return path


def parse_numstat(path_field: str) -> tuple[str, str] | None:
    """``(old_path, new_path)`` de un campo de ruta de ``--numstat``.

    ``prefijo{old => new}sufijo`` es la forma de git para un renombrado con
    prefijo o sufijo comunes; sin llaves la ruta no cambió de nombre.
    """
    if not path_field or path_field == "-":
        return None
    match = _RENAME_BRACES_RE.search(path_field)
    if match is None:
        return path_field, path_field
    prefix = path_field[:match.start()]
    suffix = path_field[match.end():]
    return prefix + match.group(1) + suffix, prefix + match.group(2) + suffix


def original_path(new_path: str, base_ref: str, run_git: RunGit, head: str = "HEAD") -> str | None:
    """Ruta donde se añadió por primera vez el fichero en ``base_ref..HEAD``.

    ``git log --follow`` sigue la cadena de renombrados; se lee su salida en
    orden inverso y se toma el primer ``A`` o la primera ruta previa.

    ``None`` si la historia no lo muestra o si ``run_git`` falla. Un fichero
    vendorizado que nació dentro de la ruta vendorizada devuelve esa misma ruta.
    """
    try:
        raw = run_git("log", "--follow", "--name-status", "--pretty=format:",
                      f"{base_ref}..{head}", "--", new_path)
    except RuntimeError:
        return None
    entries = [line for line in raw.splitlines() if line.strip()]
    for entry in reversed(entries):
        parts = entry.split("\t")
        status = parts[0]
        if status == "A" and len(parts) == _ADDED_FIELDS:
            return _strip_diff_prefix(parts[1])
        if status[:1] in ("R", "C") and len(parts) == _RENAME_FIELDS:
            return _strip_diff_prefix(parts[1])
    return None


def origins_for(base_ref: str, run_git: RunGit, head: str = "HEAD") -> str:
    """Fichero de origen del PR: una línea por fichero vendorizado de ``base_ref...HEAD``."""
    names = run_git("diff", "--name-only", "--find-renames=50%", f"{base_ref}...{head}")
    lines = []
    for path in names.splitlines():
        if not is_under_vendored(path):
            continue
        origin = original_path(path, base_ref, run_git, head)
        lines.append(f"{path}\t{origin if origin is not None else '-'}")
    return "".join(f"{line}\n" for line in lines)


def parse_origins(text: str) -> dict[str, str | None]:
    """Mapa ``ruta-nueva -> origen`` (``None`` cuando el origen es ``-``)."""
    origins: dict[str, str | None] = {}
    for line in text.splitlines():
        path, _, origin = line.partition("\t")
        if path:
            origins[path] = None if origin in ("", "-") else origin
    return origins


def _iter_numstat(numstat: str):
    """``(old_path, new_path, lines)`` de cada entrada del numstat contable.

    Una sola lectura del ``git diff --numstat`` para toda la regla: el
    presupuesto y la cifra del árbol vendorizado salen de aquí.
    """
    for line in numstat.splitlines():
        added, _, rest = line.partition("\t")
        deleted, _, path_field = rest.partition("\t")
        if not added.isdigit() or not deleted.isdigit():
            continue
        parsed = parse_numstat(path_field)
        if parsed is None:
            continue
        old_path, new_path = (_strip_diff_prefix(p) for p in parsed)
        yield old_path, new_path, int(added) + int(deleted)


def budget_lines(numstat: str, origins: dict[str, str | None] | None = None) -> int:
    """Líneas que cuentan para el presupuesto, desde un ``git diff --numstat``.

    Sin ``origins`` cuenta todo, como ``pr-contract`` antes de #931. Con
    ``origins``, un fichero vendorizado solo cuenta si su origen es de
    autoría o si es un renombrado desde una ruta de autoría. Un fichero
    vendorizado sin entrada en ``origins`` cuenta (fallo hacia contar).
    """
    total = 0
    for old_path, new_path, lines in _iter_numstat(numstat):
        if not is_under_vendored(new_path) or origins is None or new_path not in origins:
            total += lines
            continue
        origin = origins[new_path]
        if (origin is not None and not is_under_vendored(origin)) or counts_to_budget(old_path, new_path):
            total += lines
    return total


def vendored_lines(numstat: str) -> int:
    """Líneas del numstat cuyo destino cae bajo una raíz vendorizada.

    El gate la usa para nombrar la parte de la cifra que NO es trabajo de
    quien abre la PR: sin mapa de orígenes esas líneas cuentan hacia dentro
    del presupuesto (#931) y el hallazgo tiene que decirlo (#1133).
    """
    return sum(lines for _, new_path, lines in _iter_numstat(numstat) if is_under_vendored(new_path))
