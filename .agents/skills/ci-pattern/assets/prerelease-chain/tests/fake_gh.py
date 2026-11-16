#!/usr/bin/env python3
# ci-pattern asset tests — fake gh (#1627 slice 7). Registra cada llamada y responde
# con la primera regla cuyo `match` aparece en la línea de argumentos. Sin red.
import json
import os
import sys

argv = sys.argv[1:]
line = " ".join(argv)
log = os.environ.get("FAKE_GH_LOG")
if log:
    with open(log, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(argv) + "\n")
with open(os.environ["FAKE_GH_RULES"], encoding="utf-8") as fh:
    rules = json.load(fh)
for rule in rules:
    if rule["match"] in line:
        sys.stdout.write(rule.get("stdout", ""))
        sys.stderr.write(rule.get("stderr", ""))
        sys.exit(rule.get("rc", 0))
sys.stderr.write(f"fake gh: sin regla para: {line}\n")
sys.exit(97)
