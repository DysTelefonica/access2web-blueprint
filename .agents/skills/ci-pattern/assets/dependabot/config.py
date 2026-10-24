#!/usr/bin/env python3
"""Dirección de Dependabot del consumer: un ecosystem por manifiesto (#1182).

El canon publica la FORMA del `.github/dependabot.yml`: un ecosystem por
manifiesto, con el ``directory`` de cada uno donde vive ese manifiesto. El
consumer declara esos directorios en su contrato (``dependabot.ecosystems``) y
el canon genera el fichero (:func:`render`) o dice que no coincide
(:func:`findings`).

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

#: Datos canónicos por ecosystem: día, hora y prefijo del mensaje de commit.
ECOSYSTEM_DEFAULTS = {
    "pip": {"day": "monday", "time": "06:00", "prefix": "chore(platform)"},
    "github-actions": {"day": "monday", "time": "06:30", "prefix": "ci(platform)"},
    "npm": {"day": "monday", "time": "06:45", "prefix": "chore(e2e)"},
}
#: Dónde se busca el manifiesto de cada ecosystem dentro de su directorio.
MANIFEST_CANDIDATES = {
    "pip": ("pyproject.toml", "requirements*.txt"),
    "npm": ("package.json",),
    "github-actions": (".github/workflows",),
}

_ECOSYSTEM_RE = re.compile(r"(?m)^\s*-?\s*package-ecosystem:\s*[\"']?(?P<ecosystem>[A-Za-z0-9_-]+)[\"']?\s*$")
_DIRECTORY_RE = re.compile(r"(?m)^\s*directory:\s*(?P<directory>\S+)\s*$")


def declared_ecosystems(declaration: object) -> dict[str, str]:
    """Ecosystems declarados → directorio, en orden de declaración.

    Acepta ``{"ecosystems": {"pip": "app"}}`` y la forma con objeto; lo que no
    se entiende devuelve vacío (HR-3).
    """
    if not isinstance(declaration, dict):
        return {}
    ecosystems = declaration.get("ecosystems", declaration)
    if not isinstance(ecosystems, dict):
        return {}
    out: dict[str, str] = {}
    for ecosystem, value in ecosystems.items():
        if not isinstance(ecosystem, str) or not ecosystem:
            continue
        if isinstance(value, str):
            out[ecosystem] = value
        elif isinstance(value, dict) and isinstance(value.get("directory"), str):
            out[ecosystem] = value["directory"]
    return out


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


def ecosystem_block(ecosystem: str, directory: str) -> str:
    """El bloque `updates:` de un ecosystem, con los datos canónicos."""
    defaults = ECOSYSTEM_DEFAULTS.get(ecosystem, {"day": "monday", "time": "06:00", "prefix": "chore(platform)"})
    labels = "\n".join(f"      - {label}" for label in ("dependencies", "type:chore"))
    return (
        f"  - package-ecosystem: {ecosystem}\n"
        f"    directory: {directory_value(directory)}\n"
        "    schedule:\n"
        "      interval: weekly\n"
        f"      day: {defaults['day']}\n"
        f'      time: "{defaults["time"]}"\n'
        "      timezone: Etc/UTC\n"
        "    open-pull-requests-limit: 5\n"
        "    labels:\n"
        f"{labels}\n"
        "    commit-message:\n"
        f"      prefix: {defaults['prefix']}\n"
    )


def render(declaration: object) -> str:
    """El `.github/dependabot.yml` que el canon genera desde la declaración."""
    template = (Path(__file__).resolve().parents[1] / TEMPLATE_RELATIVE).read_text(encoding="utf-8")
    blocks = [ecosystem_block(ecosystem, directory)
              for ecosystem, directory in declared_ecosystems(declaration).items()]
    return template.replace(PLACEHOLDER, "".join(blocks))


def directories_in(text: str) -> dict[str, str]:
    """Ecosystem → directory de un fichero de Dependabot, leído sin parser YAML."""
    found: dict[str, str] = {}
    current = ""
    for line in text.splitlines():
        match = _ECOSYSTEM_RE.search(line)
        if match:
            current = match.group("ecosystem")
            continue
        match = _DIRECTORY_RE.search(line)
        if match and current:
            found[current] = match.group("directory")
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
    ecosystems = declared_ecosystems(declaration)
    if not ecosystems:
        return ["HR-75: la declaración 'dependabot' del contrato de adopción no declara ningún ecosystem con su directorio; declare al menos uno (ecosistema: directorio del manifiesto)"]
    out: list[str] = []
    for ecosystem, directory in ecosystems.items():
        if not manifest_exists(repo, ecosystem, directory):
            out.append(f"HR-75: el ecosystem '{ecosystem}' declara el directorio '{directory}' y ahí no hay manifiesto ({', '.join(MANIFEST_CANDIDATES.get(ecosystem, ()))})")
    consumer_file = Path(repo) / CONSUMER_PATH
    if not consumer_file.is_file():
        out.append(f"HR-75: falta {CONSUMER_PATH} en el consumer; genérelo con el canon (plantilla {TEMPLATE_RELATIVE.as_posix()}) desde la declaración 'dependabot' del contrato")
        return out
    current = consumer_file.read_text(encoding="utf-8")
    declared_dirs = {e: directory_value(d) for e, d in ecosystems.items()}
    out += _directory_findings(directories_in(current), declared_dirs)
    if not out and current != render(declaration):
        out.append(f"HR-75: {CONSUMER_PATH} no coincide con el que genera el canon desde la declaración; regenérelo (los ecosystems y sus directorios son correctos: difiere el resto del fichero)")
    return out
