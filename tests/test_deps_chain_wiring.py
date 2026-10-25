"""La cadena de dependencias (HR-76) corre igual en CI y en el preflight local.

El gate `local-ci-parity` del canon compara el comando exacto del step de
`pr-gate.yml` con su entrada en `scripts/local-preflight.sh`. Si el texto
difiere, la fase 4 de `adoption check` queda en rojo. Este test fija lo mismo
en el suite de BP: una sola línea, idéntica en los dos sitios.
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GATE = "check_deps_chain.py"


def _ci_run() -> str:
    text = (ROOT / ".github" / "workflows" / "pr-gate.yml").read_text(encoding="utf-8")
    start = text.index("      - name: Dependency chain (HR-76)")
    block = text[start:text.index("      - name: pr size", start)]
    runs = [line.strip() for line in block.splitlines() if line.strip().startswith("run:")]
    assert len(runs) == 1, runs
    return runs[0].removeprefix("run:").strip()


def _local_line() -> str:
    lines = [line.strip() for line in (ROOT / "scripts" / "local-preflight.sh").read_text(encoding="utf-8").splitlines()]
    matches = [line for line in lines if GATE in line]
    assert len(matches) == 1, matches
    return matches[0]


def test_ci_step_is_one_line_and_uses_the_gate() -> None:
    run = _ci_run()
    assert "\n" not in run
    assert GATE in run and "--changed-files" in run


def test_local_preflight_runs_the_same_command_as_ci() -> None:
    assert _local_line() == _ci_run()


def test_only_the_base_ref_variable_is_used() -> None:
    """La base es la del PR en CI y main en local; el fichero temporal sale de RUNNER_TEMP o /tmp."""
    run = _ci_run()
    assert set(re.findall(r"\$\{([A-Z_]+)", run)) == {"GITHUB_BASE_REF", "RUNNER_TEMP"}
