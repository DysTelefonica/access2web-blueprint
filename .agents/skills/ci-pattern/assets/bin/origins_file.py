#!/usr/bin/env python3
"""Escribe el fichero de origenes del PR: que fichero vendorizado llega desde donde.

El gate `pr-contract` mide el presupuesto de revision sobre el rango del PR. Sin
el mapa de origenes cuenta TODO, incluido el arbol vendorizado (`.agents/skills/`)
que el propagador materializa byte a byte (#931). El mapa dice que fichero
vendorizado llega desde una ruta de autoria (un renombrado) y cual no; solo los
primeros cuentan.

Este asset NO ejecuta git: lo ejecuta el workflow y le pasa la salida (el
ratchet extendido de los assets vendorizables no admite `subprocess` en el
asset). El mapa se arma con la misma regla del canon
(`vendored_budget.origins_for`).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ASSETS / "budget"))

from vendored_budget import origins_for  # noqa: E402


def _log_for(log: str, path: str) -> str:
    """Del log completo del rango, las entradas del fichero (y las cabeceras)."""
    out = []
    for line in log.splitlines():
        if "\t" not in line:
            out.append(line)
            continue
        if path in line.split("\t")[1:]:
            out.append(line)
    return "\n".join(out)


def _canned_run_git(names: str, log: str):
    """Un `run_git` que responde con la salida ya capturada (sin ejecutar git)."""
    def run_git(*args: str) -> str:
        if args and args[0] == "diff":
            return names
        if args and args[0] == "log":
            return _log_for(log, args[-1])
        return ""
    return run_git


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Escribe el mapa de origenes del PR (ficheros vendorizados -> su origen).")
    parser.add_argument("--base", required=True,
                        help="Ref base del PR; el rango es BASE...HEAD (merge-base, #1013).")
    parser.add_argument("--head", default="HEAD", help="Ref del PR (por defecto HEAD).")
    parser.add_argument("--names-file", required=True,
                        help="Salida de `git diff --name-only --find-renames=50%% BASE...HEAD`.")
    parser.add_argument("--log-file", required=True,
                        help="Salida de `git log --name-status --pretty=format: BASE..HEAD`.")
    parser.add_argument("--output", required=True, help="Fichero donde se escribe el mapa.")
    args = parser.parse_args(argv)
    names = Path(args.names_file).read_text(encoding="utf-8")
    log = Path(args.log_file).read_text(encoding="utf-8")
    text = origins_for(args.base, _canned_run_git(names, log), args.head)
    Path(args.output).write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
