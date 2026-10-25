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
    "github-actions": (".github/workflows", "action.yml"),
}

_ECOSYSTEM_RE = re.compile(r"(?m)^\s*-?\s*package-ecosystem:\s*[\"']?(?P<ecosystem>[A-Za-z0-9_-]+)[\"']?\s*$")
_DIRECTORY_RE = re.compile(r"(?m)^\s*directory:\s*(?P<value>\S.*?)\s*$")
_DIRECTORIES_RE = re.compile(r"(?m)^\s*directories:\s*$")
_DIRECTORIES_FLOW_RE = re.compile(r"(?m)^\s*directories:\s*\[(?P<items>[^\]]*)\]\s*(?:#.*)?$")
_LIST_ITEM_RE = re.compile(r"^\s*-\s*(?P<value>\S.*?)\s*$")


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


#: Longitud mínima de un valor entre comillas: la pareja de comillas.
MIN_QUOTED_LENGTH = 2


def _unquote(value: str) -> str:
    """Quita las comillas simples o dobles que Dependabot admite alrededor de un valor."""
    value = value.strip()
    if len(value) >= MIN_QUOTED_LENGTH and value[0] == value[-1] and value[0] in "\"'":
        return value[1:-1]
    return value


def _strip_comment(raw: str) -> str:
    """El valor de una linea de Dependabot, sin el comentario que lo cierra.

    Un `#` dentro de un valor entrecomillado NO abre comentario: `"/app#1"` es el
    directorio `/app#1` (#1221).
    """
    value = raw.strip()
    if value[:1] in ("'", '"'):
        quote = value[0]
        end = value.find(quote, 1)
        return value[: end + 1] if end != -1 else value
    index = value.find("#")
    return value[:index].strip() if index != -1 else value


def _flow_items(raw: str) -> list[str]:
    """Los items de una lista en forma flow: `[a, b]` -> `["a", "b"]`."""
    items: list[str] = []
    current = ""
    quote = ""
    for char in raw:
        if quote:
            current += char
            if char == quote:
                quote = ""
            continue
        if char in ("'", '"'):
            quote = char
            current += char
            continue
        if char == ",":
            items.append(current)
            current = ""
            continue
        current += char
    if current.strip():
        items.append(current)
    return [_strip_comment(item) for item in items if item.strip()]


def escapes_root(directory: str) -> bool:
    """True si el directorio sale del arbol del consumer.

    Un `..` en algun segmento es traversal: `/app/../..` sale del repo. Un punto en
    el nombre no lo es: `/app/..data` es un directorio DENTRO del repo (#1221).
    """
    return any(segment == ".." for segment in directory.replace("\\", "/").split("/"))


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
    for ecosystem, block in _block_lines(text).items():
        values = _block_directories(block)
        if values:
            found[ecosystem] = _unquote(values[0])
    return found


def expand_directories(repo: Path, directory: str) -> list[str]:
    """Los directorios que cubre un `directory` declarado, con sus globs expandidos.

    Un directorio sin glob es una ruta literal. Un glob (`*`, `?`) se expande contra
    el arbol del consumer: solo cuentan los directorios que existen (fail closed: un
    glob sin ningun directorio devuelve lista vacia y el gate lo nombra).
    """
    value = directory_value(directory)
    if not any(char in value for char in "*?"):
        return [value]
    pattern = value.strip("/")
    expanded = []
    for match in sorted(Path(repo).glob(pattern)):
        if match.is_dir():
            expanded.append(directory_value(match.relative_to(repo).as_posix()))
    return expanded


def _block_lines(text: str) -> dict[str, list[str]]:
    """Las lineas de cada bloque `package-ecosystem:` del fichero, sin parser YAML."""
    blocks: dict[str, list[str]] = {}
    current = ""
    for line in text.splitlines():
        match = _ECOSYSTEM_RE.search(line)
        if match:
            current = match.group("ecosystem")
            blocks.setdefault(current, [])
            continue
        if current:
            blocks[current].append(line)
    return blocks


def _block_directories(lines: list[str]) -> list[str]:
    """Los directorios declarados en un bloque de ecosystem, en orden de aparicion.

    Lee `directory:` y `directories:`, en forma de bloque o en forma flow
    (`directories: [a, b]`), y quita el comentario que cierra cada valor (#1221).
    """
    found: list[str] = []
    for index, line in enumerate(lines):
        match = _DIRECTORY_RE.search(line)
        if match:
            found.append(_strip_comment(match.group("value")))
            continue
        flow = _DIRECTORIES_FLOW_RE.search(line)
        if flow:
            found += _flow_items(flow.group("items"))
            continue
        if _DIRECTORIES_RE.search(line):
            for follow in lines[index + 1:]:
                if follow.strip().startswith("#"):
                    continue  # un comentario dentro de la lista no es un item
                item = _LIST_ITEM_RE.search(follow)
                if not item:
                    break
                found.append(_strip_comment(item.group("value")))
    return found


def coverage_in(text: str) -> dict[str, list[str]]:
    """Ecosystem -> todos sus directorios (`directory` o `directories`), sin comillas.

    Los globs se devuelven tal cual: :func:`findings` los expande contra el arbol.
    """
    found: dict[str, list[str]] = {}
    for ecosystem, block in _block_lines(text).items():
        found[ecosystem] = [directory_value(_unquote(value)) for value in _block_directories(block)]
    return found


def _groups_block(lines: list[str]) -> list[str]:
    """El sub-bloque `groups:` de un ecosystem (hasta la siguiente clave del bloque).

    El grupo se busca DENTRO de ese bloque: una clave con el mismo nombre en otro
    sitio del ecosystem (el `labels:` que escribe el render) no es el grupo
    declarado (#1221).
    """
    for index, line in enumerate(lines):
        match = re.match(r"^(?P<indent>\s*)groups:\s*$", line)
        if not match:
            continue
        indent = len(match.group("indent"))
        block: list[str] = []
        for follow in lines[index + 1:]:
            if follow.strip() and (len(follow) - len(follow.lstrip())) <= indent:
                break
            block.append(follow)
        return block
    return []


def _extra_findings(text: str, specs: dict[str, dict]) -> list[str]:
    """Cada extra declarada (grupo, `ignore`, `target-branch`) aparece en el bloque de su ecosystem."""
    out: list[str] = []
    blocks = _block_lines(text)
    for ecosystem, spec in specs.items():
        block = "\n".join(blocks.get(ecosystem, []))
        groups = "\n".join(_groups_block(blocks.get(ecosystem, [])))
        for group in spec.get("groups", {}):
            if not re.search(rf"^\s*{re.escape(group)}:\s*$", groups, re.M):
                out.append(f"HR-75: {CONSUMER_PATH} no declara el grupo '{group}' del ecosystem '{ecosystem}' que la declaración exige")
        for item in spec.get("ignore", []):
            name = str(item.get("dependency-name", ""))
            if name and not re.search(rf"dependency-name:\s*[\"']?{re.escape(name)}[\"']?\s*$", block, re.M):
                out.append(f"HR-75: {CONSUMER_PATH} no ignora '{name}' en el ecosystem '{ecosystem}' que la declaración exige")
        if "target-branch" in spec and not re.search(
                rf"target-branch:\s*[\"']?{re.escape(str(spec['target-branch']))}[\"']?\s*$", block, re.M):
            out.append(f"HR-75: {CONSUMER_PATH} no fija target-branch '{spec['target-branch']}' en el ecosystem '{ecosystem}' que la declaración exige")
    return out


def _declared_directories(repo: Path, specs: dict[str, dict]) -> tuple[dict[str, list[str]], list[str]]:
    """Directorios declarados por ecosystem (globs expandidos) y los hallazgos de declararlos."""
    declared: dict[str, list[str]] = {}
    out: list[str] = []
    for ecosystem, spec in specs.items():
        declared[ecosystem] = []
        for directory in spec["directories"]:
            if escapes_root(directory):
                out.append(f"HR-75: el ecosystem '{ecosystem}' declara el directorio '{directory}' que sale del árbol del consumer (`..`): un `directory` de Dependabot apunta dentro del repo, no a su padre")
                continue
            expanded = expand_directories(repo, directory)
            if not expanded:
                out.append(f"HR-75: el ecosystem '{ecosystem}' declara el directorio '{directory}' y el glob no expande a ningún directorio del repo")
            for item in expanded:
                declared[ecosystem].append(item)
                if not manifest_exists(repo, ecosystem, item):
                    out.append(f"HR-75: el ecosystem '{ecosystem}' declara el directorio '{item}' y ahí no hay manifiesto ({', '.join(MANIFEST_CANDIDATES.get(ecosystem, ()))})")
    return declared, out


def _coverage_findings(repo: Path, specs: dict[str, dict], declared: dict[str, list[str]], text: str) -> list[str]:
    """Ecosystems y directorios declarados que el fichero del consumer no cubre."""
    out: list[str] = []
    covered_raw = coverage_in(text)
    for ecosystem, raw_dirs in covered_raw.items():
        for raw in raw_dirs:
            if escapes_root(raw):
                out.append(f"HR-75: {CONSUMER_PATH} apunta el ecosystem '{ecosystem}' al directorio '{raw}' que sale del árbol del consumer (`..`)")
    for ecosystem in covered_raw:
        if ecosystem not in specs:
            out.append(f"HR-75: {CONSUMER_PATH} declara el ecosystem '{ecosystem}' y la declaración no lo pide; la dirección de Dependabot es dato del canon, no del consumer")
    for ecosystem, wanted in declared.items():
        if ecosystem not in covered_raw:
            out.append(f"HR-75: {CONSUMER_PATH} no declara el ecosystem '{ecosystem}' que la declaración exige")
            continue
        covered: set[str] = set()
        for raw in covered_raw[ecosystem]:
            covered.update(expand_directories(repo, raw) if any(c in raw for c in "*?") else [raw])
        out += [f"HR-75: {CONSUMER_PATH} no cubre el directorio '{directory}' del ecosystem '{ecosystem}' que la declaración exige"
                for directory in wanted if directory not in covered]
    return out


def findings(repo: Path, declaration: object) -> list[str]:
    """Hallazgos del fichero de Dependabot del consumer frente a la declaración.

    Cobertura, no igualdad (TS#1207): cada ecosystem y cada directorio declarados
    (globs expandidos) deben estar cubiertos por el fichero del consumer, con sus
    extras declaradas. Un ecosystem que nadie declaró sigue siendo un hallazgo.
    Lo que el fichero añade (labels, schedule, commit-message...) no se mide.
    """
    specs = ecosystem_specs(declaration)
    if not specs:
        return ["HR-75: la declaración 'dependabot' del contrato de adopción no declara ningún ecosystem con su directorio; declare al menos uno (ecosistema: directorio del manifiesto)"]
    out: list[str] = [f"HR-75: la declaración 'dependabot' tiene la clave '{key}' en el ecosystem '{name}', que el canon no admite (admitidas: {', '.join(sorted(ALLOWED_KEYS))})"
                      for name, key in unknown_keys(declaration)]
    declared, declared_out = _declared_directories(repo, specs)
    out += declared_out
    consumer_file = Path(repo) / CONSUMER_PATH
    if not consumer_file.is_file():
        out.append(f"HR-75: falta {CONSUMER_PATH} en el consumer; genérelo con el canon (plantilla {TEMPLATE_RELATIVE.as_posix()}) desde la declaración 'dependabot' del contrato")
        return out
    text = consumer_file.read_text(encoding="utf-8")
    out += _coverage_findings(repo, specs, declared, text)
    out += _extra_findings(text, specs)
    return out
