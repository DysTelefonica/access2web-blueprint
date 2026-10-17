#!/usr/bin/env python3
# check_docs_architecture.py — gate del documento de arquitectura (HR-68, #714)
"""``docs/architecture.md`` no puede mentir: árbol↔código y capas declaradas.

- **Árbol ↔ código, en los dos sentidos**: cada módulo de primer nivel descrito
  en la tabla de «Árbol de módulos» existe bajo algún code-root, y cada módulo
  que existe bajo algún code-root está descrito. Un lado vacío es un FAIL.
- **Capas**: el bloque de datos ``architecture`` (JSON, bloque cercado) declara
  las capas con sus módulos y las dependencias permitidas. Todo módulo del árbol
  pertenece a exactamente una capa, toda capa tiene su regla y toda regla
  nombra capas declaradas. Lo que no se permite explícitamente está prohibido.

Stdlib exclusivamente, sin red; mismo script en local y CI. Exit 0 limpio,
1 hallazgos, 2 entrada inválida. Las reglas de imports las impone el adaptador
del linter del stack leyendo este mismo bloque (``docs-architecture`` PR B).
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from pathlib import Path

BLOCK_INFO = "architecture"
TREE_HEADING = "árbol de módulos"
_HEADING_RE = re.compile(r"^#{1,6}\s+(?:\d+\.\s+)?(.*?)\s*#*\s*$")
_CELL_PATH_RE = re.compile(r"^`([^`]+)`$")


def _load_fences():
    """Carga ``assets/markdown_fences.py`` por ruta, igual que ``docs-codeguide``."""
    path = Path(__file__).resolve().parent.parent / "markdown_fences.py"
    spec = importlib.util.spec_from_file_location("markdown_fences", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_FENCES = _load_fences()


def _load_paths():
    """Carga ``architecture_paths.py`` (capas por patrón de ruta) hermano, por ruta."""
    path = Path(__file__).resolve().parent / "architecture_paths.py"
    spec = importlib.util.spec_from_file_location("architecture_paths", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _load_imports():
    """Carga el adaptador de imports ``architecture_imports.py`` hermano, por ruta."""
    path = Path(__file__).resolve().parent / "architecture_imports.py"
    spec = importlib.util.spec_from_file_location("architecture_imports", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def list_module_paths(repo: Path, code_root: str, code_glob: str) -> list[str]:
    """Módulos de primer nivel bajo ``code_root`` como rutas relativas al repo.

    Misma regla que ``docs-codeguide``: un fichero que casa el glob, o un
    directorio con al menos un fichero que lo casa. ``__init__.py`` no es un
    módulo, es el marcador del paquete."""
    root = repo / code_root
    if not root.is_dir():
        return []
    found = []
    for entry in sorted(root.iterdir()):
        if entry.name == "__init__.py":
            continue
        if (entry.is_file() and entry.match(code_glob)) or (
            entry.is_dir() and any(entry.glob(code_glob))
        ):
            found.append(f"{code_root.rstrip('/')}/{entry.name}")
    return found


def parse_tree(text: str) -> list[str]:
    """Rutas de la primera columna de las tablas bajo «Árbol de módulos».

    Los subtítulos más profundos pertenecen al árbol; la sección acaba en el
    siguiente título de su nivel o superior."""
    tree_level, paths = 0, []
    for _number, line in _FENCES.unfenced_lines(text):
        heading = _HEADING_RE.match(line)
        if heading:
            level = len(line) - len(line.lstrip("#"))
            if heading.group(1).casefold() == TREE_HEADING:
                tree_level = level
            elif tree_level and level <= tree_level:
                tree_level = 0
            continue
        if tree_level and line.lstrip().startswith("|"):
            cell = _CELL_PATH_RE.match(line.strip().strip("|").split("|")[0].strip())
            if cell:
                paths.append(cell.group(1).rstrip("/"))
    return paths


def parse_block(text: str):
    """``(datos, error)`` del bloque de capas.

    Es el primer bloque cercado con información ``architecture`` o, si no hay, el primer
    bloque cercado entre las marcas ``<!-- architecture:begin -->`` y ``<!-- architecture:end -->``."""
    lines = text.splitlines()
    begin = next((n for n, line in enumerate(lines, 1) if line.strip() == "<!-- architecture:begin -->"), 0)
    end = next((n for n, line in enumerate(lines, 1) if line.strip() == "<!-- architecture:end -->"), 0)
    blocks = list(_FENCES.fenced_blocks(text))
    chosen = next((b for b in blocks if b[0] == BLOCK_INFO), None) or next(
        (b for b in blocks if begin and end and begin < b[2] < end), None)
    if chosen is None:
        return None, f"falta el bloque cercado «{BLOCK_INFO}» con las capas y las dependencias"
    try:
        data = json.loads(chosen[1])
    except json.JSONDecodeError as exc:
        return None, f"el bloque «{BLOCK_INFO}» no es JSON válido: {exc}"
    if not isinstance(data, dict):
        return None, f"el bloque «{BLOCK_INFO}» debe ser un objeto"
    return data, None


def _layer_findings(data: dict, tree: list[str] | None) -> list[str]:
    layers, allowed = data.get("layers"), data.get("allowed")
    if not isinstance(layers, list) or not layers or not isinstance(allowed, dict):
        return ["el bloque necesita «layers» (lista no vacía) y «allowed» (objeto)"]
    findings, names, owner = [], [], {}
    for layer in layers:
        name = layer.get("name") if isinstance(layer, dict) else None
        modules = layer.get("modules") if isinstance(layer, dict) else None
        if not isinstance(name, str) or not name or not isinstance(modules, list) or not modules:
            findings.append(f"capa mal formada: {layer!r} (necesita name y modules no vacío)")
            continue
        if name in names:
            findings.append(f"capa duplicada: {name}")
        names.append(name)
        for module in modules:
            if tree is None:
                continue
            if module not in tree:
                findings.append(f"la capa {name} declara {module} y no está en el árbol")
            elif module in owner:
                findings.append(f"el módulo {module} pertenece a varias capas: {owner[module]}, {name}")
            else:
                owner[module] = name
    findings += [f"el módulo {m} no pertenece a ninguna capa declarada" for m in tree or [] if m not in owner]
    for name in names:
        if name not in allowed:
            findings.append(f"la capa {name} no tiene regla en «allowed» (use [] si no puede depender de ninguna)")
    for source, targets in allowed.items():
        if source not in names:
            findings.append(f"«allowed» nombra la capa {source}, que no está declarada")
        for target in targets if isinstance(targets, list) else [None]:
            if target not in names:
                findings.append(f"«allowed» permite {source} → {target}, capa no declarada")
    return findings + _rule_findings(data, names, allowed)


def _rule_findings(data: dict, names: list, allowed: dict) -> list[str]:
    """Claves opcionales del bloque (#749): ``independent``, ``pure`` y ``tolerated``."""
    findings = [f"«{key}» debe ser una lista" for key in ("independent", "tolerated")
                if key in data and not isinstance(data[key], list)]
    if findings:
        return findings
    for group in data.get("independent", []):
        mods = group.get("modules") if isinstance(group, dict) else None
        if not isinstance(mods, list) or len(mods) < 2 or not all(isinstance(m, str) and m for m in mods):
            findings.append(f"«independent» necesita grupos con al menos dos módulos: {group!r}")
    pure = data.get("pure")
    if pure is not None:
        layers, externals = (pure.get("layers"), pure.get("forbidden_externals")) if isinstance(pure, dict) else (None, None)
        if not isinstance(layers, list) or not layers or not isinstance(externals, list) or not externals:
            findings.append("«pure» necesita «layers» y «forbidden_externals» no vacíos")
        else:
            findings += [f"«pure» nombra la capa {n}, que no está declarada" for n in layers if n not in names]
    for item in data.get("tolerated", []):
        src, dst, edges = (item.get("from"), item.get("to"), item.get("imports")) if isinstance(item, dict) else (None, None, None)
        if src not in names or dst not in names:
            findings.append(f"«tolerated» nombra una capa no declarada: {src} → {dst}")
        elif src == dst or dst in allowed.get(src, []):
            findings.append(f"«tolerated» {src} → {dst} no tiene sentido: la dependencia ya está permitida")
        if not isinstance(edges, list) or not edges or not all(isinstance(e, str) and e.count("->") == 1 for e in edges):
            findings.append(f"«tolerated» {src} → {dst} necesita «imports» no vacío con la forma «origen -> destino»")
    return findings


def _load_imports_ts():
    """Carga el adaptador TypeScript ``architecture_imports_ts.py`` hermano, por ruta."""
    path = Path(__file__).resolve().parent / "architecture_imports_ts.py"
    spec = importlib.util.spec_from_file_location("architecture_imports_ts", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


STACKS = ("python", "typescript")


def _adapter(stack: str):
    return _load_imports_ts() if stack == "typescript" else _load_imports()


def _resolved_block(repo: Path, text: str, tree: list[str], roots: list[tuple[str, str]]):
    """``(bloque con modules por capa, hallazgos)``; los patrones de ruta se expanden a paquetes."""
    data, error = parse_block(text)
    if error:
        return None, [error]
    if any(isinstance(layer, dict) and "paths" in layer for layer in data.get("layers", [])):
        data, path_findings = _load_paths().normalize(repo, data, roots)
        return data, path_findings + _layer_findings(data, None)
    return data, _layer_findings(data, tree)


def check_architecture(repo: Path, doc: str, roots: list[tuple[str, str]],
                       linter_config: str | None = None, stack: str = "python") -> list[str]:
    """Hallazgos del documento ``doc``; lista vacía = verde. Falla cerrado.

    Con ``linter_config`` valida además las reglas de imports del ``stack``
    (``python`` con import-linter, ``typescript`` con dependency-cruiser):
    la configuración versionada coincide con la derivada del bloque y el
    linter se ejecuta sin violaciones."""
    if not repo.is_dir():
        raise FileNotFoundError(f"repo path is not a directory: {repo}")
    if not roots:
        raise ValueError("at least one (code_root, code_glob) pair is required")
    page = repo / doc
    if not page.is_file():
        return [f"falta el documento {doc}"]
    text = page.read_text(encoding="utf-8")
    tree = parse_tree(text)
    findings = []
    if not tree:
        findings.append(f"<empty árbol de módulos en {doc}>")
    actual: list[str] = []
    for code_root, code_glob in roots:
        modules = list_module_paths(repo, code_root, code_glob)
        if not modules:
            findings.append(f"<empty {code_root}/>")
        actual += modules
    findings += [f"árbol → código: {doc} describe {m} y no existe" for m in tree if not (repo / m).exists()]
    findings += [f"código → árbol: el módulo {m} no está descrito en {doc}"
                 for m in actual if not any(t == m or t.startswith(m + "/") for t in tree)]
    data, layer_findings = _resolved_block(repo, text, tree, roots)
    findings += layer_findings
    if linter_config and not layer_findings:
        findings += _adapter(stack).check_imports(repo, data, linter_config)
    return findings


def write_config(repo: Path, doc: str, linter_config: str, stack: str = "python",
                 roots: list[tuple[str, str]] | None = None) -> list[str]:
    """Escribe la configuración del linter derivada del bloque; devuelve los hallazgos si no es válido.

    Con capas por patrón de ruta necesita los code-roots para expandirlas a paquetes."""
    text = (repo / doc).read_text(encoding="utf-8") if (repo / doc).is_file() else ""
    data, problems = _resolved_block(repo, text, parse_tree(text), roots or [])
    if not problems:
        render = _load_imports_ts().render_depcruise if stack == "typescript" else _load_imports().render_importlinter
        (repo / linter_config).write_text(render(data), encoding="utf-8")
    return problems


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="check_docs_architecture.py",
        description="Comprueba docs/architecture.md contra el código (HR-68, #714).",
    )
    parser.add_argument("repo", help="Ruta al directorio raíz del repo a validar.")
    parser.add_argument("--doc", default="docs/architecture.md", help="Documento de arquitectura.")
    parser.add_argument("--code-root", action="append", default=None,
                        help="Raíz de código (repetible, emparejada con --code-glob).")
    parser.add_argument("--code-glob", action="append", default=None,
                        help="Glob que un módulo debe contener para contar (repetible).")
    parser.add_argument("--stack", choices=STACKS, default="python",
                        help="Stack del linter de imports (python: import-linter; typescript: dependency-cruiser).")
    parser.add_argument("--linter-config", default=None,
                        help="Configuración del linter de imports a validar.")
    parser.add_argument("--write-config", action="store_true",
                        help="Escribe --linter-config derivada del bloque y sale.")
    return parser


def _pin_output_encoding() -> None:
    """Hace la salida del gate independiente de la locale del runner."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str]) -> int:
    _pin_output_encoding()
    args = _build_parser().parse_args(argv[1:])
    if not args.code_root or not args.code_glob or len(args.code_root) != len(args.code_glob):
        print("FAIL input: --code-root y --code-glob son obligatorios y van por pares", file=sys.stderr)
        return 2
    roots = list(zip(args.code_root, args.code_glob, strict=True))
    if args.write_config:
        problems = write_config(Path(args.repo), args.doc, args.linter_config or ".importlinter",
                                args.stack, roots)
        for problem in problems:
            print(f"FAIL architecture: {problem}", file=sys.stderr)
        return 1 if problems else 0
    try:
        findings = check_architecture(Path(args.repo), args.doc, roots, args.linter_config, args.stack)
    except FileNotFoundError as exc:
        print(f"FAIL input: {exc}", file=sys.stderr)
        return 2
    for finding in findings:
        print(f"FAIL architecture: {finding}", file=sys.stderr)
    if findings:
        return 1
    print(f"docs-architecture: OK ({args.doc}, {len(roots)} code-root(s))")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
