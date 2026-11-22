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
import json
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


def _yaml_code(text: str) -> str:
    """Quita los comentarios YAML: un '#' al inicio de linea o tras blanco abre
    comentario. Un comentario no es un uso real del vocabulario de liveness."""
    out = []
    for line in text.splitlines():
        m = re.search(r"(^|\s)#", line)
        out.append(line[:m.start()] if m else line)
    return "\n".join(out)


def _extract_jobs(content: str) -> list[str]:
    """Ids de job de PRIMER nivel: las claves a la indentacion del primer job
    bajo la clave `jobs:` de columna 0. Las claves anidadas (name, run, env...)
    no son gates."""
    jobs, in_jobs, job_indent = [], False, None
    for line in content.splitlines():
        if re.match(r"^jobs\s*:", line):
            in_jobs, job_indent = True, None
            continue
        if not in_jobs:
            continue
        if line.strip() and not line[0].isspace():
            in_jobs = False
            continue
        m = re.match(r"^(\s+)(\w[\w-]*):", line)
        if not m:
            continue
        indent = len(m.group(1).expandtabs(2))
        if job_indent is None:
            job_indent = indent
        if indent == job_indent:
            jobs.append(m.group(2))
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
        if any(kw in _yaml_code(wf.read_text(encoding="utf-8")).lower()
               for kw in _KW_LIV):
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


def _required_jobs(repo: Path) -> tuple[list[str], str | None]:
    """Gates declarados en `.github/required-jobs-policy.json` (`required_jobs`).
    Devuelve (nombres, error): error a None si el repo no declara politica o si
    la politica es legible; si no, el motivo por el que no se puede leer."""
    path = repo / ".github" / "required-jobs-policy.json"
    if not path.is_file():
        return [], None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return [], f"no se puede leer {path.name}: {exc}"
    names = data.get("required_jobs") if isinstance(data, dict) else None
    if not isinstance(names, list) or not all(isinstance(n, str) for n in names):
        return [], f"{path.name}: required_jobs debe ser una lista de nombres de job"
    return list(names), None


def _defined_jobs(repo: Path) -> set[str]:
    out: set[str] = set()
    for wf in _find_yamls(repo):
        out.update(_extract_jobs(wf.read_text(encoding="utf-8")))
    return out


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

    required, policy_error = _required_jobs(a.repo)
    if policy_error is not None:
        print(f"INVALID: política required_jobs no legible en {a.repo}: {policy_error}")
        return 2
    defined = _defined_jobs(a.repo) if required else set()
    uninvoked = [n for n in required if n not in defined]

    gates = _detect_gates(a.repo)
    if not gates and not uninvoked:
        print(f"SKIP: no se detectaron gates fail-closed en {a.repo}")
        return 2

    if uninvoked:
        print(f"FINDING: {len(uninvoked)} gate(s) declarado(s) en required_jobs "
              "que ningún workflow invoca.")
        for name in uninvoked:
            print(f"  - {name} (required_jobs) sin job en .github/workflows")
        return 1

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
