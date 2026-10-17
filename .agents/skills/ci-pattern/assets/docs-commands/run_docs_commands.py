#!/usr/bin/env python3
# run_docs_commands.py — entrada del job `docs-commands`: lee la declaración y ejecuta el gate.
"""``run_docs_commands.py <repo> [--config docs-commands.json]``: exit 0/1/2."""

import argparse
import importlib.util
import json
import sys
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument("repo")
parser.add_argument("--config", default="docs-commands.json", help="JSON con las claves de docs_commands.")
args = parser.parse_args()
spec = importlib.util.spec_from_file_location("check_docs_commands", Path(__file__).with_name("check_docs_commands.py"))
gate = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gate)
try:
    config = json.loads((Path(args.repo) / args.config).read_text(encoding="utf-8"))
except (OSError, json.JSONDecodeError) as exc:
    print(f"ERROR: configuración ilegible: {exc}", file=sys.stderr)
    sys.exit(2)
findings, summary = gate.check(args.repo, config)
print("docs-commands: " + ", ".join(f"{tier}={count}" for tier, count in summary.items()))
print("\n".join(findings))
sys.exit(1 if findings else 0)
