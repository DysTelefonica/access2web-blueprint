#!/usr/bin/env python3
"""Dirección de Dependabot del consumer: un ecosystem por manifiesto (#1182, #1193).

El canon publica la FORMA del `.github/dependabot.yml`: un ecosystem por
manifiesto, con el ``directory`` de cada uno donde vive ese manifiesto. El
consumer declara esos directorios en su contrato (``dependabot.ecosystems``) y
el canon genera el fichero (:func:`render`) o dice que no coincide
(:func:`findings`).

Cada ecosystem admite, además del directorio, las claves ``directories``,
``target-branch``, ``ignore`` y ``groups`` (#1193). Cualquier otra clave es un
hallazgo (HR-3): el canon no descarta configuración en silencio.

Por qué es dato del canon y no una copia por consumer: Dependabot actualiza el
**manifiesto**, no el lock; el lock lo regenera el PR que cambia el manifiesto, y
un lock atrás hace que el bump no se aplique (#1181). No ejecuta ``git`` ni lee
la red: lee el árbol del consumer y el contrato.
"""

from __future__ import annotations

import fnmatch
import re
from pathlib import Path

#: Clave de la declaración en el contrato de adopción.
DECLARATION_KEY = "dependabot"

#: Dónde vive el fichero del consumer.
CONSUMER_PATH = ".github/dependabot.yml"

#: La plantilla canónica, con el bloque `updates` como único VALOR.
TEMPLATE_RELATIVE = Path("templates/.github/dependabot.yml.tmpl")
PLACEHOLDER = "{{ECOSYSTEMS}}"

#: Claves que una entrada de ecosystem admite además de su directorio (#1193).
ALLOWED_KEYS = frozenset({"directory", "directories", "target-branch", "ignore", "groups"})

#: Datos canónicos por ecosystem: día, hora y prefijo del mensaje de commit.
ECOSYSTEM_DEFAULTS = {
    "pip": {"day": "monday", "time": "06:00", "prefix": "chore(platform)"},
    "github-actions": {"day": "monday", "time": "06:30", "prefix": "ci(platform)"},
    "npm": {"day": "monday", "time": "06:45", "prefix": "chore(e2e)"},
}
DEFAULT_DEFAULTS = {"day": "monday", "time": "06:00", "prefix": "chore(platform)"}

#: Dónde se busca el manifiesto de cada ecosystem dentro de su directorio.
MANIFEST_CANDIDATES = {
    "pip": ("pyproject.toml", "requirements*.txt"),
    "npm": ("package.json",),
    "github-actions": (".github/workflows",),
}

_ECOSYSTEM_RE = re.compile(r"(?m)^\s*-?\s*package-ecosystem:\s*[\"']?(?P<ecosystem>[A-Za-z0-9_-]+)[\"']?\s*$")
_DIRECTORY_RE = re.compile(r"(?m)^\s*directory:\s*(?P<directory>\S+)\s*$")
_DIRECTORIES_RE = re.compile(r"(?m)^\s*directories:\s*$")
_LIST_ITEM_RE = re.compile(r"^\s*-\s*[\"']?(?P<item>[^\"'\s]+)[\"']?\s*$")


def _entries(declaration: object) -> dict[str, object]:
    """Entradas crudas ``ecosystem -> valor``, en orden de declaración."""
    if not isinstance(declaration, dict):
        return {}
    ecosystems = declaration.get("ecosystems", declaration)
    if not isinstance(ecosystems, dict):
        return {}
    return {name: value for name, value in ecosystems.items() if isinstance(name, str) and name}


def ecosystem_specs(declaration: object) -> dict[str, dict]:
    """Ecosystem → especificación normalizada: ``directories`` (lista) y claves extra.

    La forma escalar (``"app"``) equivale a ``{"directory": "app"}``. Una entrada
    que no se entiende no aparece: :func:`findings` lo reporta (HR-3).
    """
    out: dict[str, dict] = {}
    for name, value in _entries(declaration).items():
        if isinstance(value, str):
            out[name] = {"directories": [value]}
            continue
        if not isinstance(value, dict):
            continue
        if isinstance(value.get("directories"), list) and value["directories"] \
                and all(isinstance(item, str) for item in value["directories"]):
            directories = list(value["directories"])
        elif isinstance(value.get("directory"), str):
            directories = [value["directory"]]
        else:
            continue
        spec = {key: val for key, val in value.items() if key in ALLOWED_KEYS and key not in ("directory", "directories")}
        spec["directories"] = directories
        out[name] = spec
    return out


def unknown_keys(declaration: object) -> list[tuple[str, str]]:
    """Pares ``(ecosystem, clave)`` que la declaración escribe y el canon no conoce."""
    found: list[tuple[str, str]] = []
    for name, value in _entries(declaration).items():
        if isinstance(value, dict):
            found += [(name, key) for key in value if key not in ALLOWED_KEYS]
    return found


def declared_ecosystems(declaration: object) -> dict[str, str]:
    """Ecosystems declarados → directorio principal (el primero), en orden de declaración."""
    return {name: spec["directories"][0] for name, spec in ecosystem_specs(declaration).items()}


def directory_value(directory: str) -> str:
    """El `directory` como lo escribe Dependabot: con barra inicial (la raíz, `/`)."""
    cleaned = directory.strip().strip("/")
    return "/" + cleaned if cleaned not in ("", ".") else "/"


def manifest_exists(repo: Path, ecosystem: str, directory: str) -> bool:
    """True cuando en el directorio declarado hay un manifiesto de ese ecosystem.

    Un ecosystem sin candidatos conocidos se acepta tal cual.
    """
    candidates = MANIFEST_CANDIDATES.get(ecosystem)
    if not candidates:
        return True
    base = Path(repo) / directory.strip("/") if directory.strip("/") else Path(repo)
    if not base.is_dir():
        return False
    for candidate in candidates:
        if any(char in candidate for char in "*?"):
            if any(fnmatch.fnmatch(entry.name, candidate)
                   for entry in base.iterdir() if entry.is_file()):
                return True
        elif (base / candidate).exists():
            return True
    return False


def _quoted(value: str) -> str:
    return '"' + value.replace('"', '\\"') + '"'


def _groups_lines(groups: dict) -> list[str]:
    """``groups:`` con cada grupo y sus claves (``patterns`` o ``dependency-type``)."""
    out = ["    groups:"]
    for name, body in groups.items():
        out.append(f"      {name}:")
        for key, val in body.items():
            if isinstance(val, list):
                out.append(f"        {key}:")
                out += [f"          - {_quoted(item)}" for item in val]
            else:
                out.append(f"        {key}: {_quoted(str(val))}")
    return out


def ecosystem_block(ecosystem: str, spec: dict) -> str:
    """El bloque `updates:` de un ecosystem: datos canónicos más las claves declaradas."""
    defaults = ECOSYSTEM_DEFAULTS.get(ecosystem, DEFAULT_DEFAULTS)
    directories = [directory_value(item) for item in spec["directories"]]
    lines = [f"  - package-ecosystem: {ecosystem}"]
    if len(directories) == 1:
        lines.append(f"    directory: {directories[0]}")
    else:
        lines.append("    directories:")
        lines += [f"      - {_quoted(item)}" for item in directories]
    if "target-branch" in spec:
        lines.append(f"    target-branch: {_quoted(str(spec['target-branch']))}")
    lines += [
        "    schedule:",
        "      interval: weekly",
        f"      day: {defaults['day']}",
        f'      time: "{defaults["time"]}"',
        "      timezone: Etc/UTC",
        "    open-pull-requests-limit: 5",
    ]
    if "ignore" in spec:
        lines.append("    ignore:")
        lines += [f"      - dependency-name: {_quoted(str(item['dependency-name']))}" for item in spec["ignore"]]
    if "groups" in spec:
        lines += _groups_lines(spec["groups"])
    lines += ["    labels:", "      - dependencies", "      - type:chore"]
    lines += ["    commit-message:", f"      prefix: {defaults['prefix']}"]
    return "\n".join(lines) + "\n"


def render(declaration: object) -> str:
    """El `.github/dependabot.yml` que el canon genera desde la declaración."""
    template = (Path(__file__).resolve().parents[1] / TEMPLATE_RELATIVE).read_text(encoding="utf-8")
    blocks = [ecosystem_block(ecosystem, spec) for ecosystem, spec in ecosystem_specs(declaration).items()]
    return template.replace(PLACEHOLDER, "\n".join(blocks))


def directories_in(text: str) -> dict[str, str]:
    """Ecosystem → directorio principal de un fichero de Dependabot, sin parser YAML."""
    found: dict[str, str] = {}
    current = ""
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = _ECOSYSTEM_RE.search(line)
        if match:
            current = match.group("ecosystem")
            continue
        if not current:
            continue
        match = _DIRECTORY_RE.search(line)
        if match:
            found[current] = match.group("directory")
            current = ""
            continue
        if _DIRECTORIES_RE.search(line):
            for follow in lines[index + 1:]:
                item = _LIST_ITEM_RE.search(follow)
                if not item:
                    break
                found.setdefault(current, item.group("item"))
            current = ""
    return found


def _directory_findings(found_dirs: dict[str, str], declared_dirs: dict[str, str]) -> list[str]:
    """Los ecosystems cuya dirección no es la declarada, y los que nadie declaró."""
    out: list[str] = []
    for ecosystem, expected in declared_dirs.items():
        got = found_dirs.get(ecosystem)
        if got is None:
            out.append(f"HR-75: {CONSUMER_PATH} no declara el ecosystem '{ecosystem}' que la declaración exige con el directorio '{expected}'")
        elif got != expected:
            out.append(f"HR-75: {CONSUMER_PATH} apunta el ecosystem '{ecosystem}' al directorio '{got}' y la declaración exige '{expected}' (el manifiesto de ese ecosystem vive en '{expected}')")
    for ecosystem in found_dirs:
        if ecosystem not in declared_dirs:
            out.append(f"HR-75: {CONSUMER_PATH} declara el ecosystem '{ecosystem}' y la declaración no lo pide; la dirección de Dependabot es dato del canon, no del consumer")
    return out


def findings(repo: Path, declaration: object) -> list[str]:
    """Hallazgos del fichero de Dependabot del consumer frente a la declaración."""
    specs = ecosystem_specs(declaration)
    if not specs:
        return ["HR-75: la declaración 'dependabot' del contrato de adopción no declara ningún ecosystem con su directorio; declare al menos uno (ecosistema: directorio del manifiesto)"]
    out: list[str] = [f"HR-75: la declaración 'dependabot' tiene la clave '{key}' en el ecosystem '{name}', que el canon no admite (admitidas: {', '.join(sorted(ALLOWED_KEYS))})"
                      for name, key in unknown_keys(declaration)]
    for ecosystem, spec in specs.items():
        for directory in spec["directories"]:
            if not manifest_exists(repo, ecosystem, directory):
                out.append(f"HR-75: el ecosystem '{ecosystem}' declara el directorio '{directory}' y ahí no hay manifiesto ({', '.join(MANIFEST_CANDIDATES.get(ecosystem, ()))})")
    consumer_file = Path(repo) / CONSUMER_PATH
    if not consumer_file.is_file():
        out.append(f"HR-75: falta {CONSUMER_PATH} en el consumer; genérelo con el canon (plantilla {TEMPLATE_RELATIVE.as_posix()}) desde la declaración 'dependabot' del contrato")
        return out
    current = consumer_file.read_text(encoding="utf-8")
    out += _directory_findings(directories_in(current), declared_dirs_of(specs))
    if not out and current != render(declaration):
        out.append(f"HR-75: {CONSUMER_PATH} no coincide con el que genera el canon desde la declaración; regenérelo (los ecosystems y sus directorios son correctos: difiere el resto del fichero)")
    return out


def declared_dirs_of(specs: dict[str, dict]) -> dict[str, str]:
    """Directorio principal por ecosystem, normalizado como lo escribe Dependabot."""
    return {name: directory_value(spec["directories"][0]) for name, spec in specs.items()}
