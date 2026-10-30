#!/usr/bin/env python3
# run_docs_links.py — entrada del job `docs-links`: lee la declaración y ejecuta el gate.
"""``run_docs_links.py <repo> [--contract .github/ci-pattern-adoption.json]``: exit 0/1/2.

La fuente de la declaración es la clave ``docs_links`` del contrato de
adopción (#1020). ``--config`` sigue aceptando el fichero de raíz antiguo,
avisa por stderr de que está obsoleto y no lo cambia de alcance.
"""

import argparse
import importlib.util
import json
import sys
from pathlib import Path

CONTRACT = ".github/ci-pattern-adoption.json"

parser = argparse.ArgumentParser()
parser.add_argument("repo")
parser.add_argument("--contract", default=CONTRACT,
                    help="Contrato de adopción; de ahí se lee la clave docs_links.")
parser.add_argument("--config",
                    help="OBSOLETO (#1020): JSON suelto con las claves de docs_links. "
                         "La fuente única es " + CONTRACT + ".")
args = parser.parse_args()
spec = importlib.util.spec_from_file_location("check_docs_links", Path(__file__).with_name("check_docs_links.py"))
gate = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gate)
try:
    if args.config:
        print(f"AVISO: --config está obsoleto (#1020); la fuente única de docs_links es "
              f"{CONTRACT}. Bórrelo cuando el job deje de invocarlo.", file=sys.stderr)
        config = json.loads((Path(args.repo) / args.config).read_text(encoding="utf-8"))
    else:
        contract = json.loads((Path(args.repo) / args.contract).read_text(encoding="utf-8"))
        config = contract["docs_links"]
except (OSError, json.JSONDecodeError, KeyError, TypeError) as exc:
    print(f"ERROR: no se pudo leer la declaración de docs_links desde "
          f"{args.config or args.contract}: {exc}", file=sys.stderr)
    sys.exit(2)
if not isinstance(config, dict):
    print(f"ERROR: la declaración de docs_links no es un objeto; revise "
          f"{args.config or args.contract}", file=sys.stderr)
    sys.exit(2)
findings = gate.check(args.repo, config)
print("\n".join(findings) if findings else "docs-links: OK")
sys.exit(1 if findings else 0)
