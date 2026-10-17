#!/usr/bin/env python3
# ci-pattern asset — gate de perfil de runners (DysTelefonica/team-skills#435)
"""check_runner_profile.py <repo>: gobierna el perfil de runners declarado
en el ci-pattern.yaml del consumer (P51–P53).

Regla (HR-61): con `P51_runners_profile: self-hosted`, todo job de un
workflow disparado por `pull_request` o `push` debe correr en el pool
declarado por `P52_runners_pool_labels`, o ser una excepción hosted
justificada en `P53_hosted_exceptions` (job → motivo no vacío, estilo
HR-15). Sin perfil declarado el default es `hosted` y el gate es inocuo
(retrocompatibilidad). Un repo privado Free sin minutos hosted que no
declara el perfil bloquea su propio CI: eso es exactamente lo que este
gate hace visible.

Fail-closed (HR-3):
- ci-pattern.yaml ausente o ilegible → exit 2.
- `P51` con valor fuera de {hosted, self-hosted} → exit 2.
- perfil self-hosted sin `P52_runners_pool_labels` no vacía → exit 2.
- `runs-on` dudoso (dinámico, mapa sin `labels`, o ausente heredando el
  default hosted) es hallazgo, no silencio.

Hallazgos (exit 1), una línea por caso:
- HOSTED-SIN-EXCEPCION  job hosted sin entrada en P53.
- MOTIVO-VACIO          excepción declarada sin motivo validable.
- FUERA-DE-POOL         runs-on self-hosted con etiquetas distintas de P52.
- SIN-RUNS-ON           job sin runs-on: hereda hosted.
- RUNS-ON-DINAMICO      runs-on con expresión ${{ }} no evaluable.
- RUNS-ON-DUDOSO        valor no evaluable por otras causas y sin excepción.

Salida: stdout con veredictos machine-legibles; stderr errores de entrada.
Exit codes: 0 limpio/inocuo · 1 hallazgos · 2 entrada inválida.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

HOSTED_IMAGE = re.compile(r"^(ubuntu|windows|macos)(-[\w.@+-]+)?$")


class EntradaInvalida(Exception):
    """El gate no puede interpretar el entorno; sale 2, nunca adivina."""


# ---------------------------------------------------------------------------
# Subconjunto YAML (mismo contrato que parse_params_yaml de la CLI: mapas
# anidados por indentación de 2, listas en bloque o inline, sin tabs).
# ---------------------------------------------------------------------------


def _strip_comment(line: str) -> str:
    out = []
    quote = ""
    for ch in line:
        if quote:
            out.append(ch)
            if ch == quote:
                quote = ""
        elif ch in "\"'":
            quote = ch
            out.append(ch)
        elif ch == "#":
            break
        else:
            out.append(ch)
    return "".join(out)


def _parse_scalar(text: str):
    text = text.strip()
    if not text:
        return None
    if text.startswith("[") and text.endswith("]"):
        inner = text[1:-1].strip()
        if not inner:
            return []
        return [_unquote(part.strip()) for part in inner.split(",")]
    return _unquote(text)


def _unquote(text: str):
    if len(text) >= 2 and text[0] == text[-1] and text[0] in "\"'":
        return text[1:-1]
    if text in ("true", "false"):
        return text
    return text


def parse_yaml_subset(text: str) -> dict:
    root: dict = {}
    stack = [(-1, None, None)]
    for lineno, raw in enumerate(text.splitlines(), start=1):
        if "\t" in raw:
            raise EntradaInvalida(f"línea {lineno}: tabulación prohibida")
        line = _strip_comment(raw).rstrip()
        if not line.strip():
            continue
        indent = len(line) - len(line.lstrip(" "))
        stripped = line.strip()
        while len(stack) > 1 and indent <= stack[-1][0]:
            stack.pop()
        _ind, parent, key = stack[-1]
        container = root if key is None else parent[key]
        if stripped.startswith("- "):
            if key is None or not isinstance(container, (list, type(None))):
                raise EntradaInvalida(f"línea {lineno}: ítem de lista mal posicionado")
            if container is None:
                container = parent[key] = []
            container.append(_parse_scalar(stripped[2:]))
            continue
        if ":" not in stripped:
            raise EntradaInvalida(f"línea {lineno}: falta ':' en '{stripped}'")
        new_key, _, rest = stripped.partition(":")
        new_key = new_key.strip().strip("\"'")
        if isinstance(container, list):
            raise EntradaInvalida(f"línea {lineno}: mapa dentro de lista no soportado")
        if container is None:
            container = parent[key] = {}
        if not isinstance(container, dict):
            raise EntradaInvalida(
                f"línea {lineno}: '{key}' tiene valor escalar; no admite subclaves")
        if new_key in container:
            raise EntradaInvalida(f"línea {lineno}: clave duplicada '{new_key}'")
        container[new_key] = _parse_scalar(rest)
        stack.append((indent, container, new_key))
    return root


# ---------------------------------------------------------------------------
# Parametros del consumer
# ---------------------------------------------------------------------------


def load_params(repo: Path) -> dict:
    path = repo / "ci-pattern.yaml"
    if not path.is_file():
        raise EntradaInvalida(
            f"ci-pattern.yaml ausente en {path}; sin contrato no hay perfil gobernable")
    try:
        return parse_yaml_subset(path.read_text(encoding="utf-8"))
    except EntradaInvalida as exc:
        raise EntradaInvalida(f"ci-pattern.yaml ilegible ({path}): {exc}") from exc


def profile_of(params: dict):
    raw = params.get("P51_runners_profile")
    if raw is None:
        return "hosted"
    value = str(raw)
    if value not in ("hosted", "self-hosted"):
        raise EntradaInvalida(
            f"P51_runners_profile={value!r} fuera de {{hosted, self-hosted}}")
    return value


def pool_of(params: dict):
    labels = params.get("P52_runners_pool_labels")
    if labels is None:
        labels = []
    if isinstance(labels, str):
        labels = [labels]
    return [str(x) for x in labels]


def exceptions_of(params: dict) -> dict:
    raw = params.get("P53_hosted_exceptions")
    return raw if isinstance(raw, dict) else {}


# ---------------------------------------------------------------------------
# Workflows
# ---------------------------------------------------------------------------


def has_pr_or_push_trigger(doc: dict) -> bool:
    trig = doc.get("on", doc.get(True))
    if trig is None:
        return False
    if isinstance(trig, str):
        return trig in ("pull_request", "push")
    if isinstance(trig, list):
        return any(str(x) in ("pull_request", "push") for x in trig)
    if isinstance(trig, dict):
        return "pull_request" in trig or "push" in trig
    return False


def classify_runs_on(value, pool: set) -> tuple:
    """(clase, detalle). clase ∈ pool | hosted | fuera-de-pool | dinamico | dudoso."""
    if isinstance(value, dict):
        value = value.get("labels")
        if value is None:
            return "dudoso", "runs-on mapa sin labels"
    if isinstance(value, list):
        labels = [str(x) for x in value]
        if "self-hosted" in labels:
            if set(labels) == pool:
                return "pool", ",".join(labels)
            return "fuera-de-pool", f"self-hosted con etiquetas {labels} != pool {sorted(pool)}"
        return "dudoso", f"lista sin self-hosted: {labels}"
    text = str(value or "")
    if text.startswith("${{"):
        return "dinamico", f"runs-on dinámico {text}"
    if HOSTED_IMAGE.match(text):
        return "hosted", text
    if text == "self-hosted":
        if pool == {"self-hosted"}:
            return "pool", text
        return "fuera-de-pool", "self-hosted sin las etiquetas del pool"
    return "dudoso", f"runs-on no evaluable: {text!r}"


def gate_workflow(path: Path, pool_labels, exceptions) -> list:
    doc = parse_yaml_subset(path.read_text(encoding="utf-8"))
    if not has_pr_or_push_trigger(doc):
        return []
    pool = set(pool_labels)
    defaults = (doc.get("defaults") or {}).get("jobs") or {}
    default_runs_on = defaults.get("runs-on")
    findings = []
    jobs = doc.get("jobs")
    if not isinstance(jobs, dict):
        return findings
    for job_id, job in jobs.items():
        if not isinstance(job, dict):
            continue
        runs_on = job.get("runs-on", job.get("runs_on"))
        if runs_on is None:
            runs_on = default_runs_on
        if runs_on is None:
            findings.append(
                f"SIN-RUNS-ON: {path.name} :: job '{job_id}' sin runs-on "
                "hereda el default hosted")
            continue
        clase, detalle = classify_runs_on(runs_on, pool)
        exc_declared = job_id in exceptions
        if exc_declared and not str(exceptions.get(job_id) or "").strip():
            findings.append(
                f"MOTIVO-VACIO: {path.name} :: job '{job_id}' excepción "
                "sin motivo validable (HR-15)")
            continue
        if clase == "hosted" and not exc_declared:
            findings.append(
                f"HOSTED-SIN-EXCEPCION: {path.name} :: job '{job_id}' "
                f"pide '{detalle}' sin entrada en P53_hosted_exceptions")
        elif clase == "fuera-de-pool":
            findings.append(f"FUERA-DE-POOL: {path.name} :: job '{job_id}' — {detalle}")
        elif clase == "dinamico":
            findings.append(f"RUNS-ON-DINAMICO: {path.name} :: job '{job_id}' — {detalle}")
        elif clase == "dudoso":
            findings.append(f"RUNS-ON-DUDOSO: {path.name} :: job '{job_id}' — {detalle}")
    return findings


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv) -> int:
    _pin_output_encoding()
    repo_arg = argv[1] if len(argv) > 1 else "."
    repo = Path(repo_arg)
    if not repo.is_dir():
        print(f"ERROR: repo inexistente: {repo}", file=sys.stderr)
        return 2
    try:
        params = load_params(repo)
        profile = profile_of(params)
        if profile == "hosted":
            print("PERFIL: hosted — gate inocuo (default P51 no declarado)")
            return 0
        pool_labels = pool_of(params)
        if not pool_labels:
            raise EntradaInvalida(
                "perfil self-hosted sin P52_runners_pool_labels: no hay pool que gobernar")
        wf_dir = repo / ".github" / "workflows"
        files = sorted(list(wf_dir.glob("*.yml")) + list(wf_dir.glob("*.yaml"))) if wf_dir.is_dir() else []
        exceptions = exceptions_of(params)
        findings = []
        governed = 0
        for wf in files:
            try:
                own = gate_workflow(wf, pool_labels, exceptions)
            except EntradaInvalida as exc:
                raise EntradaInvalida(f"{wf}: {exc}") from exc
            governed += 1
            findings.extend(own)
        print(f"PERFIL: self-hosted | pool: {pool_labels} | "
              f"workflows gobernados: {governed} | hallazgos: {len(findings)}")
        for line in findings:
            print(line)
        return 1 if findings else 0
    except EntradaInvalida as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
