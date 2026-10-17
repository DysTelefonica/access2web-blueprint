#!/usr/bin/env python3
# run_docs_links.py — entrada del job `docs-links`: lee la declaración y ejecuta el gate.
"""``run_docs_links.py <repo> [--config docs-links.json]``: exit 0/1/2."""

import argparse
import importlib.util
import json
import sys
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument("repo")
parser.add_argument("--config", default="docs-links.json", help="JSON con las claves de docs_links.")
args = parser.parse_args()
spec = importlib.util.spec_from_file_location("check_docs_links", Path(__file__).with_name("check_docs_links.py"))
gate = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gate)
try:
    config = json.loads((Path(args.repo) / args.config).read_text(encoding="utf-8"))
except (OSError, json.JSONDecodeError) as exc:
    print(f"ERROR: configuración ilegible: {exc}", file=sys.stderr)
    sys.exit(2)
findings = gate.check(args.repo, config)
print("\n".join(findings) if findings else "docs-links: OK")
sys.exit(1 if findings else 0)
