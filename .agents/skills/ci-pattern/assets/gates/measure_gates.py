#!/usr/bin/env python3
"""Línea base de fricción de gates (#1179, fase 0).

Sobre un EXPORTE de check-runs (JSON, sin red) calcula las cifras que deciden
si el CI de un repo publica un check por gate o uno solo: **rojos
secundarios** (por commit, los checks rojos que no son el primero: la
redundancia), **gates rojos por commit** (el peor caso), **commits con
re-ejecución** (verdes sin commit nuevo) y **tiempo hasta el verde** (la
latencia de la CI, que la auditoría de #1093 no midió).

Exclusiones (las de la auditoría): el agregador (`required`, que repite el rojo
de la hoja) y los `cancelled`/`skipped` (no son fallos). Y se mide la PRIMERA
onda de la CI: una re-ejecución posterior no es la espera del merge.

Exit: 0 métricas · 2 export inválido o sin commits (HR-3).
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

#: El agregador de required checks repite el rojo de la hoja que lo causa.
AGGREGATE_NAMES = frozenset({"required"})
AGGREGATE_SUFFIX = " (required)"
#: Conclusiones que SÍ son un fallo del gate.
FAILURE_CONCLUSIONS = frozenset({"failure", "timed_out", "startup_failure"})
#: Percentil del tiempo hasta el verde que se reporta (además de la mediana).
P90 = 0.90
#: Ventana de la primera onda de la CI, en minutos: un run que empieza más
#: tarde pertenece a otra onda (re-ejecución, o job filtrado por paths que
#: solo corrió cuando el PR se actualizó) y no cuenta como espera del merge.
WAVE_WINDOW_MINUTES = 15
#: Los mensajes de error viven como constantes de módulo: el `raise` lleva el
#: nombre, no el texto (así el lint no dispara TRY003 en el call site).
ERR_UNREADABLE = "export ilegible: {detail}"
ERR_SHAPE = "el export debe ser un objeto con 'commits' (lista)"
ERR_EMPTY = "el export no tiene commits: no hay nada que medir"
ERR_COMMIT = "cada commit del export debe ser un objeto"
ERR_NO_VALUES = "no hay valores para el percentil"


class InvalidExportError(ValueError):
    """El export no se puede interpretar: el asset sale 2 (HR-3)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def _parse_time(value) -> datetime | None:
    """ISO-8601 a datetime con zona horaria; `None` si no se puede leer."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def first_wave(checks: list) -> list:
    """El primer run de cada check dentro de la primera onda de la CI.

    Una re-ejecución es otro run del mismo nombre: para medir la latencia de
    la CI cuenta la primera onda (la que el merge espera), no la última.
    """
    starts = [s for s in (_parse_time(run.get("started_at")) for run in checks) if s]
    if not starts:
        return []
    limit = min(starts) + timedelta(minutes=WAVE_WINDOW_MINUTES)
    wave: dict[str, dict] = {}
    for run in checks:
        started = _parse_time(run.get("started_at"))
        if started is None or started > limit:
            continue
        name = str(run.get("name") or "")
        current = wave.get(name)
        if current is None or started < _parse_time(current.get("started_at")):
            wave[name] = run
    return list(wave.values())


def failing_checks(checks: list) -> list:
    """Checks en rojo de la primera onda, sin el agregador ni los no-fallos."""
    red = []
    for check in checks:
        name = str(check.get("name") or "")
        if str(check.get("conclusion") or "") not in FAILURE_CONCLUSIONS:
            continue
        if name in AGGREGATE_NAMES or name.endswith(AGGREGATE_SUFFIX):
            continue
        red.append(name)
    return red


def has_rerun(checks: list) -> bool:
    """¿El commit tuvo algún check que corriera más de una vez?"""
    names = [str(run.get("name") or "") for run in checks]
    return len(names) > len(set(names))


def green_delay_minutes(checks: list) -> float | None:
    """Minutos de la latencia de la CI: del primer run al último verde."""
    wave = first_wave(checks)
    starts = [s for s in (_parse_time(run.get("started_at")) for run in wave) if s]
    finished = [e for e in (_parse_time(run.get("completed_at")) for run in wave
                            if str(run.get("conclusion") or "") == "success") if e]
    if not starts or not finished:
        return None
    return (max(finished) - min(starts)).total_seconds() / 60.0


def percentile(values: list, fraction: float) -> float:
    """Percentil por interpolación del vecino más cercano."""
    if not values:
        raise InvalidExportError(ERR_NO_VALUES)
    ordered = sorted(values)
    position = min(len(ordered) - 1, max(0, round(fraction * (len(ordered) - 1))))
    return ordered[position]


def load_export(path: str) -> dict:
    """Lee y valida el export; cualquier forma inválida sale 2."""
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise InvalidExportError(ERR_UNREADABLE.format(detail=exc)) from exc
    if not isinstance(doc, dict) or not isinstance(doc.get("commits"), list):
        raise InvalidExportError(ERR_SHAPE)
    if not doc["commits"]:
        raise InvalidExportError(ERR_EMPTY)
    return doc


def measure(doc: dict) -> dict:
    """Calcula las métricas del export; el dict es el veredicto."""
    red_commits = secondary = with_checks = reruns = without_green = 0
    worst = {"pr": None, "sha": None, "reds": 0, "checks": []}
    delays = []
    for commit in doc["commits"]:
        if not isinstance(commit, dict):
            raise InvalidExportError(ERR_COMMIT)
        checks = commit.get("checks") or []
        if not checks:
            continue
        with_checks += 1
        reruns += 1 if has_rerun(checks) else 0
        red = failing_checks(first_wave(checks))
        if red:
            red_commits += 1
            secondary += len(red) - 1
            if len(red) > worst["reds"]:
                worst = {"pr": commit.get("pr"), "sha": commit.get("sha"),
                         "reds": len(red), "checks": sorted(red)}
        delay = green_delay_minutes(checks)
        without_green += 1 if delay is None else 0
        if delay is not None:
            delays.append(delay)

    def ratio(part: int) -> float:
        """Porción de los commits con checks (el denominador de cada %)."""
        return round(part / with_checks, 3) if with_checks else 0

    green = {"measured": len(delays),
             "median": round(percentile(delays, 0.5), 1) if delays else None,
             "p90": round(percentile(delays, P90), 1) if delays else None,
             "max": round(max(delays), 1) if delays else None}
    return {"commits": len(doc["commits"]), "commits_with_checks": with_checks,
            "red_commits": red_commits, "red_ratio": ratio(red_commits),
            "secondary_reds": secondary, "worst_commit": worst,
            "commits_with_rerun": reruns, "rerun_ratio": ratio(reruns),
            "commits_without_green": without_green, "time_to_green_minutes": green}


def render(doc: dict, metrics: dict) -> str:
    """Veredicto legible: una línea por métrica, con su denominador."""
    window = doc.get("window") or {}
    green = metrics["time_to_green_minutes"]
    worst = metrics["worst_commit"]
    return "\n".join([
        f"repo: {doc.get('repo') or 'sin nombre'}",
        f"ventana: {window.get('from') or '?'} → {window.get('to') or '?'}",
        f"commits medidos: {metrics['commits']} "
        f"(con checks: {metrics['commits_with_checks']})",
        f"commits con rojo: {metrics['red_commits']} ({metrics['red_ratio'] * 100:.1f} %)",
        f"rojos secundarios: {metrics['secondary_reds']} (los que no son el primero)",
        f"peor commit: {worst['reds']} checks rojos (PR {worst['pr']}, sha {worst['sha']})",
        f"commits con re-ejecución: {metrics['commits_with_rerun']} "
        f"({metrics['rerun_ratio'] * 100:.1f} %)",
        f"tiempo hasta el verde (min): mediana {green['median']}, p90 {green['p90']}, "
        f"máximo {green['max']} (medidos {green['measured']}, sin verde "
        f"{metrics['commits_without_green']})",
    ])


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="Métricas de fricción de gates (#1179, fase 0).")
    parser.add_argument("--export", required=True, help="export JSON de check-runs")
    parser.add_argument("--json", action="store_true", help="salida máquina")
    args = parser.parse_args(argv)
    try:
        doc = load_export(args.export)
        metrics = measure(doc)
    except InvalidExportError as exc:
        print(f"GATES-MEASURE ERROR: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(metrics, indent=2, sort_keys=True) if args.json else render(doc, metrics))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
