#!/usr/bin/env python3
"""Reparto del dominio de mutación entre los jobs de la matriz (#785).

El dominio tiene 1129 mutantes y la sesión entera no cabe en un job (38-113 h
medidas). Este script reparte los módulos de una sesión de `cosmic-ray init` en
tramos deterministas con greedy LPT —el mayor al tramo menos cargado, desempate
por índice y ruta— y escribe el config de un tramo con el `test-command`,
timeout y exclusiones del config compartido. Cada job hace su propio `init`
(3,7 s) y el job final fusiona las sesiones.

Uso:
    python scripts/mutation_shards.py --session s.sqlite --shards 10 --summary
    python scripts/mutation_shards.py --session s.sqlite --index 3 --out shard.toml
"""

from __future__ import annotations

import argparse
import importlib.machinery
import importlib.util
import json
import sys
import tomllib
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "docs" / "quality" / "cosmic-ray.toml"


class ShardError(Exception):
    """Precondición incumplida: el reparto no se puede calcular ni escribir."""


def _load_check_mutation():
    """Reutiliza la lectura de la sesión del gate: un solo lector del layout."""
    path = ROOT / "scripts" / "check_mutation.py"
    spec = importlib.util.spec_from_loader(
        "check_mutation", importlib.machinery.SourceFileLoader("check_mutation", str(path))
    )
    module = importlib.util.module_from_spec(spec)
    # Los dataclasses de `check_mutation` resuelven su módulo en `sys.modules`
    # durante el decorador: sin registrarlo antes del `exec_module` falla.
    sys.modules["check_mutation"] = module
    spec.loader.exec_module(module)
    return module


def module_counts(session: Path) -> dict[str, int]:
    """Mutantes por módulo de una sesión ya inicializada."""
    gate = _load_check_mutation()
    rows, errors = gate.read_session(session)
    if errors:
        raise ShardError(errors[0])
    return dict(Counter(row["module_path"] for row in gate.active_rows(rows)))


def assign(counts: dict[str, int], shards: int) -> list[dict]:
    """Reparto greedy LPT determinista; devuelve un tramo por índice."""
    if shards < 1:
        raise ShardError("--shards debe ser >= 1")
    if len(counts) < shards:
        raise ShardError(
            f"{len(counts)} módulo(s) para {shards} tramo(s): no se puede partir "
            "un módulo entre dos jobs"
        )
    order = sorted(counts, key=lambda module: (-counts[module], module))
    buckets: list[list[str]] = [[] for _ in range(shards)]
    loads = [0] * shards
    for module in order:
        index = min(range(shards), key=lambda i: (loads[i], i))
        buckets[index].append(module)
        loads[index] += counts[module]
    return [
        {"index": index, "mutants": loads[index], "modules": sorted(buckets[index])}
        for index in range(shards)
    ]


def render(shared: dict, modules: list[str]) -> str:
    """Config del tramo: mismo timeout, test-command y exclusiones que el compartido."""
    lines = [
        "# Config de un tramo de la matriz de mutación (#785). GENERADO por",
        "# scripts/mutation_shards.py: no lo edite a mano. El reparto se calcula",
        "# de la sesión de `cosmic-ray init`; el resto viene de",
        "# docs/quality/cosmic-ray.toml.",
        "",
        "[cosmic-ray]",
        "module-path = [",
    ]
    lines += [f'    "{module}",' for module in sorted(modules)]
    lines += ["]", f"timeout = {float(shared['timeout'])}", "excluded-modules = ["]
    lines += [f'    "{pattern}",' for pattern in shared.get("excluded-modules") or ()]
    lines += [
        "]",
        # `json.dumps` produce una cadena básica de TOML válida: el test-command
        # lleva comillas dobles dentro (`sh -c '... -m "not integration"'`) y sin
        # escapar rompería el config generado.
        f"test-command = {json.dumps(shared['test-command'])}",
        "",
        "[cosmic-ray.distributor]",
        f'name = "{shared["distributor"]["name"]}"',
        "",
    ]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--session", type=Path, required=True, help="sesión de cosmic-ray init")
    parser.add_argument("--shards", type=int, default=10, help="número de tramos (defecto 10)")
    parser.add_argument("--index", type=int, help="tramo a escribir como config")
    parser.add_argument("--out", type=Path, help="fichero de config del tramo")
    parser.add_argument("--summary", action="store_true", help="imprime la carga de cada tramo")
    args = parser.parse_args(argv)

    if (args.index is None) != (args.out is None):
        print("error: indique --index y --out juntos (o solo --summary)", file=sys.stderr)
        return 2
    if args.index is None and not args.summary:
        print("error: indique --summary, o --index y --out", file=sys.stderr)
        return 2

    try:
        counts = module_counts(args.session)
        if not counts:
            raise ShardError(
                "la sesión no tiene mutantes; `cosmic-ray init` no generó trabajo "
                "(una sesión vacía haría pasar el ratchet midiendo nada)"
            )
        plan = assign(counts, args.shards)
    except ShardError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if args.summary:
        loads = [entry["mutants"] for entry in plan]
        print(
            f"reparto: {len(counts)} módulo(s), {sum(loads)} mutante(s) en "
            f"{len(plan)} tramo(s); cargas {sorted(loads)}"
        )

    if args.index is None:
        return 0

    entries = {entry["index"]: entry for entry in plan}
    if args.index not in entries:
        print(
            f"error: no existe el tramo {args.index}; el reparto declara 0..{args.shards - 1}",
            file=sys.stderr,
        )
        return 2
    modules = entries[args.index]["modules"]
    missing = [module for module in modules if not (ROOT / module).is_file()]
    if missing:
        print(
            "error: el reparto nombra módulos que no existen en el árbol: "
            + ", ".join(sorted(missing)),
            file=sys.stderr,
        )
        return 2

    with CONFIG_PATH.open("rb") as handle:
        shared = tomllib.load(handle)["cosmic-ray"]
    args.out.write_text(render(shared, modules), encoding="utf-8")
    print(
        f"tramo {args.index}/{args.shards}: {len(modules)} módulo(s), "
        f"{entries[args.index]['mutants']} mutante(s) → {args.out}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
