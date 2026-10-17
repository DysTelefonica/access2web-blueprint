# architecture_paths.py — capas por patrón de ruta del documento de arquitectura (HR-68, #751)
"""Bloque v2: cada capa declara ``paths`` (patrones de ruta) en lugar de ``modules``.

``exempt`` gana a las capas. El primer patrón que casa con un fichero fuente decide su
capa, en el orden en que se declaran las capas; ``*`` no cruza ``/`` y ``**`` cruza directorios. ``exempt`` lista los
patrones que no pertenecen a ninguna capa. Todo fichero fuente de los code-roots cae en
una capa o en ``exempt`` (uno sin capa es rojo, HR-3) y todo patrón casa al menos un
fichero (uno que no casa nada es rojo). Para los linters, ``normalize`` expande los
patrones a los paquetes y módulos concretos y devuelve un bloque equivalente con
``modules`` por capa: el menor número de rutas que cubre exactamente los ficheros de la
capa (un directorio entero cuando todos sus ficheros son de la capa, el fichero si no).
Stdlib exclusivamente.
"""

from __future__ import annotations

import re
from pathlib import Path

_SKIP_DIRS = {".git", "node_modules", "__pycache__", ".venv"}


def glob_regex(pattern: str) -> re.Pattern:
    """``*`` no cruza ``/``, ``**`` cruza directorios, ``?`` es un carácter; el resto, literal."""
    out, i = [], 0
    while i < len(pattern):
        if pattern.startswith("**", i):
            out.append(".*")
            i += 2
        elif pattern[i] == "*":
            out.append("[^/]*")
            i += 1
        elif pattern[i] == "?":
            out.append("[^/]")
            i += 1
        else:
            out.append(re.escape(pattern[i]))
            i += 1
    return re.compile("".join(out) + r"\Z")


def _walk(repo: Path, top: str) -> list[str]:
    base = repo / top
    if not base.is_dir():
        return []
    return sorted(p.relative_to(repo).as_posix() for p in base.rglob("*")
                  if p.is_file() and not _SKIP_DIRS & set(p.relative_to(repo).parts))


def source_files(repo: Path, roots: list[tuple[str, str]]) -> list[str]:
    """Ficheros fuente: los de cada code-root, recursivos, que casan su glob."""
    found = set()
    for root, glob in roots:
        base = repo / root
        found.update(p.relative_to(repo).as_posix() for p in base.rglob(glob)
                     if p.is_file() and not _SKIP_DIRS & set(p.relative_to(repo).parts))
    return sorted(found)


def _first_layer(path: str, layers: list[tuple[str, list[re.Pattern]]]):
    for name, patterns in layers:
        if any(p.match(path) for p in patterns):
            return name
    return None


def _pack(node: str, files: dict[str, str | None], children: dict[str, list[str]], out: dict) -> None:
    """Asigna a ``out[capa]`` las rutas mínimas bajo ``node``; los exentos no cuentan."""
    if node in files:
        if files[node]:
            out.setdefault(files[node], []).append(node)
        return
    owned = {files[f] for f in files if f.startswith(node + "/") and files[f]}
    if len(owned) == 1:
        out.setdefault(owned.pop(), []).append(node)
        return
    for child in children.get(node, []):
        _pack(child, files, children, out)


def normalize(repo: Path, data: dict, roots: list[tuple[str, str]]):
    """``(bloque con modules por capa, hallazgos)`` para un bloque con ``paths``."""
    findings = []
    bad = [layer for layer in data["layers"] if not isinstance(layer, dict) or not layer.get("name")
           or not isinstance(layer.get("paths"), list) or not layer["paths"]
           or not all(isinstance(p, str) and p for p in layer["paths"])]
    if bad or not isinstance(data.get("exempt", []), list):
        return data, [f"capa mal formada: {layer!r} (con «paths» cada capa necesita name y paths no vacío)"
                      for layer in bad] or ["«exempt» debe ser una lista de patrones"]
    layers = [(layer["name"], [glob_regex(p) for p in layer["paths"]]) for layer in data["layers"]]
    exempt = [glob_regex(p) for p in data.get("exempt", [])]
    tops = sorted({root.split("/")[0] for root, _glob in roots})
    universe = [f for top in tops for f in _walk(repo, top)]
    for layer in data["layers"]:
        findings += [f"la capa {layer['name']} declara el patrón {p} y no casa ningún fichero"
                     for p in layer["paths"] if not any(glob_regex(p).match(f) for f in universe)]
    findings += [f"«exempt» declara el patrón {p} y no casa ningún fichero"
                 for p in data.get("exempt", []) if not any(glob_regex(p).match(f) for f in universe)]
    files: dict[str, str | None] = {}
    for path in source_files(repo, roots):
        layer = None if any(p.match(path) for p in exempt) else _first_layer(path, layers)
        if layer is None and not any(p.match(path) for p in exempt):
            findings.append(f"el fichero {path} no pertenece a ninguna capa ni a «exempt»")
        files[path] = layer
    children: dict[str, list[str]] = {}
    for path in files:
        parts = path.split("/")
        for depth in range(1, len(parts)):
            parent, child = "/".join(parts[:depth]), "/".join(parts[:depth + 1])
            if child not in children.setdefault(parent, []):
                children[parent].append(child)
    packed: dict[str, list[str]] = {}
    for top in tops:
        _pack(top, files, children, packed)
    allowed = data.get("allowed") or {layer["name"]: layer.get("allowed", []) for layer in data["layers"]}
    normalized = {**data, "allowed": allowed,
                  "layers": [{"name": layer["name"], "modules": packed.get(layer["name"], [])}
                             for layer in data["layers"]]}
    return normalized, findings
