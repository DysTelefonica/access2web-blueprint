#!/usr/bin/env python3
"""Liveness gate — alerta cuando un gate fail-closed se bloquea (F3, #396, HR-3/HR-21).

Un gate fail-closed que se queda parado (healthz down, e2e fallido, deploy
tumbado) no produce ninguna seÃ±al externa: el CI no avisa que el gate lleva
horas/dÃ­as bloqueado. El VPS tumbado el 2026-10-05 dejÃ³ 4 dÃ­as de producciÃ³n
caÃ­da y 74 commits mergeados sin desplegar (#1290).

Este activo verifica que el consumer tiene un mecanismo de aviso con umbral
de tiempo para gates fail-closed.

Salida:
  - OK          -> el mecanismo de aviso existe y estÃ¡ configurado.
  - FINDING     -> no hay aviso de umbral para al menos un gate fail-closed.
  - SKIP        -> no se detectaron gates fail-closed (repo sin CI deploy).

Exit: 0 OK/skip, 1 finding, 2 invalid/empty.
"""
import argparse
import re
import sys
from pathlib import Path

_KW_FCL = frozenset(["gate","block","require","verify","check","security","e2e",
    "deploy","release","publish","health","smoke","production","test"])
_KW_LIV = frozenset(["liveness","liveness_threshold","alert_after","stale_threshold",
    "gate_timeout","fail_loud_after","alert_threshold","liveness_minutes","stale_minutes",
    "gate-liveness","gate_alert","gate-stale","fail-closed-alert"])


def _find_yamls(repo: Path) -> list[Path]:
    out = []
    for d in (repo / ".github" / "workflows", repo / "ci" / "workflows",
              repo / "tools" / "ci"):
        if d.is_dir():
            out.extend(sorted(d.glob("*.yml")))
    return out


def _extract_jobs(content: str) -> list[str]:
    jobs, in_jobs = [], False
    for line in content.splitlines():
        if re.match(r"^jobs\s*:", line.strip()):
            in_jobs = True
            continue
        if in_jobs:
            if line.strip() and not line[0:1].isspace() and not line[0:1] == '\t':
                in_jobs = False
                continue
            m = re.match(r"^\s{2,}(\w[\w-]*):\s*(.*)$", line)
            if m:
                jobs.append(m.group(1))
    return jobs


def _is_fail_closed(name: str, block: str, content: str) -> bool:
    nl = name.lower()
    if any(kw in nl for kw in _KW_FCL):
        return True
    for m in re.finditer(r"needs:\s*\[([^\]]+)\]", block):
        if any(kw in m.group(1).lower() for kw in ("deploy","release","publish")):
            return True
    if re.search(r"gate_", content):
        return True
    return False


def _has_liveness(repo: Path) -> bool:
    for d in (repo / "assets" / "gate-liveness",
              repo / "personal" / "ardelperal" / "ci-pattern" / "assets" / "gate-liveness",
              repo / "tools" / "gate-liveness"):
        if d.is_dir():
            return True
    for wf in _find_yamls(repo):
        if any(kw in wf.read_text(encoding="utf-8").lower() for kw in _KW_LIV):
            return True
    for f in repo.rglob("*.json"):
        if ".git" in str(f):
            continue
        try:
            if any(kw in f.read_text(encoding="utf-8", errors="ignore").lower()
                   for kw in _KW_LIV):
                return True
        except (OSError, PermissionError):
            pass
    return False


def _detect_gates(repo: Path) -> list[dict]:
    gates = []
    for wf in _find_yamls(repo):
        content = wf.read_text(encoding="utf-8")
        rel = str(wf.relative_to(repo))
        for name in _extract_jobs(content):
            if _is_fail_closed(name, "", content):
                gates.append({"name": name, "workflow": rel})
    return gates


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    p = argparse.ArgumentParser(
        description="Liveness gate: detecta gates fail-closed sin umbral de aviso (F3, #396).")
    p.add_argument("--repo", required=True, type=Path)
    p.add_argument("--threshold", type=int, default=24)
    p.add_argument("--check", action="store_true")
    a = p.parse_args()

    gates = _detect_gates(a.repo)
    if not gates:
        print(f"SKIP: no se detectaron gates fail-closed en {a.repo}")
        return 2

    if not _has_liveness(a.repo):
        print(f"FINDING: {len(gates)} gate(s) fail-closed sin umbral de aviso "
              f"(F3, #396, HR-3).")
        for g in gates:
            print(f"  - {g['name']} en {g['workflow']}")
        print(f"  Umbral configurado: {a.threshold} horas.")
        print("  SoluciÃ³n: configurar aviso externo (issue, correo, webhook).")
        return 1

    print(f"OK: {len(gates)} gate(s) fail-closed detectados, liveness configurado.")
    for g in gates:
        print(f"  - {g['name']} en {g['workflow']} (umbral: {a.threshold}h)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
