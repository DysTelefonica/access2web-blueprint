#!/usr/bin/env python3
# ci-pattern asset — genera el fichero de anotaciones del check rojo (DysTelefonica/team-skills#1219)
"""Escribe el fichero que el gate lee con ``--annotations``: la primera anotación
de error de cada check rojo.

Las anotaciones de un check solo existen en la API de check runs: el contexto
``needs`` del runner no las trae. Este asset las lee; el gate sigue sin red.

La salida es INFORMATIVA, nunca un veredicto: si el API no se puede leer, el
asset avisa y no escribe el fichero, y el gate degrada a «no annotation». El
veredicto del agregador no depende de las anotaciones, así que un API caído no
pone el job en rojo.

El nombre con el que el runner publica cada check lo declara la política del
consumer (``check_names``), con el marcador ``{event}`` para los nombres que
dependen del evento. Un nombre declarado que no aparece entre los check runs
del SHA se avisa: la declaración del consumer está obsoleta.

Solo se piden las anotaciones de un check ROJO: un check ya terminado con ese
nombre publicado y una conclusión que no sea verde, el mismo criterio que el
lado del ``needs`` usa para decidir que un job está rojo. Un check verde no
presta su anotación a un job rojo: su texto no es la causa (#1243).

Fail-closed (HR-3): una entrada ilegible o una respuesta de API que no tiene
la forma esperada es un AVISO con el motivo, no una anotación inventada.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

try:
    from check_required_jobs import GateError, load_needs, load_policy
except ImportError:
    print("ERROR: check_required_jobs.py must be installed next to build_annotations.py",
          file=sys.stderr)
    sys.exit(2)

#: El cliente REST compartido del catálogo (runner-controls/github_rest.py), que
#: se carga por ruta como hace el gate de contrato de PR (#938).
CLIENT = Path(__file__).resolve().parent.parent / "runner-controls" / "github_rest.py"

#: Conclusiones que no son causa: el check no está rojo.
GREEN = ("success", "skipped")


def _load_client():
    """Carga `runner-controls/github_rest.py` por ruta; None si no está."""
    if not CLIENT.is_file():
        return None
    spec = importlib.util.spec_from_file_location("build_annotations_github_rest", str(CLIENT))
    if spec is None or spec.loader is None:
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def warn(message: str) -> None:
    """Un problema del generador es un AVISO: el veredicto no depende de él."""
    print(f"AVISO: {message}", file=sys.stderr)


def published_name(job: str, check_names: dict, event: str) -> str:
    """El nombre con el que el runner publica el check de este job (#1219)."""
    declared = check_names.get(job)
    if not isinstance(declared, str):
        return job
    return declared.replace("{event}", event)


def fetch_check_runs(client, repo: str, sha: str) -> list:
    """Los check runs del SHA; la API devuelve 100 por página."""
    data = client.get_json(f"repos/{repo}/commits/{sha}/check-runs?per_page=100")
    if not isinstance(data, dict) or not isinstance(data.get("check_runs"), list):
        raise TypeError("the API did not return a check-runs object")  # noqa: TRY003
    return data["check_runs"]


def find_red_run(runs: list, published: str) -> dict | None:
    """El check run ROJO con ese nombre publicado; ``None`` si no hay ninguno.

    Solo cuenta un check ya terminado con una conclusión que no sea verde
    (``GREEN``): el mismo criterio que el lado del ``needs`` usa para decidir
    que un job está rojo. Un check verde NO presta su anotación a un job rojo:
    el gate la imprimiría como si fuera la causa, y no lo es (#1243).
    """
    for run in runs:
        if not isinstance(run, dict) or run.get("name") != published:
            continue
        if run.get("status") == "completed" and run.get("conclusion") not in GREEN:
            return run
    return None


def first_annotation(client, repo: str, run: dict) -> str:
    """La primera anotación de error del check run; vacía si no tiene ninguna."""
    data = client.get_json(f"repos/{repo}/check-runs/{run['id']}/annotations?per_page=100")
    if not isinstance(data, list):
        raise TypeError("the API did not return an annotations array")  # noqa: TRY003
    entries = [item for item in data if isinstance(item, dict)]
    errors = [item for item in entries if item.get("annotation_level") == "failure"]
    for item in errors + entries:
        message = item.get("message")
        if isinstance(message, str) and message.strip():
            return message.strip()
    return ""


def write_out(path: str, data: dict) -> None:
    """El fichero que el gate lee: `job -> mensaje`, la forma que ya acepta."""
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                           encoding="utf-8")


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="build the required-jobs annotations file")
    parser.add_argument("--policy", required=True, help="path to the required-jobs policy JSON")
    parser.add_argument("--needs-file", default=None,
                        help="path to the serialized needs JSON (default: read stdin)")
    parser.add_argument("--repo", required=True, help="owner/name of the repository")
    parser.add_argument("--sha", required=True, help="head SHA whose check runs are read")
    parser.add_argument("--event", required=True, help="current event name (github.event_name)")
    parser.add_argument("--out", required=True, help="path of the annotations file to write")
    args = parser.parse_args(argv)
    try:
        policy = load_policy(args.policy)
        needs, _ = load_needs(args.needs_file)
    except (GateError, OSError) as exc:
        warn(f"no se pudo leer la politica o el needs: {exc}")
        return 0
    check_names = policy.get("check_names") or {}
    red = sorted(job for job, conclusion in needs.items() if conclusion not in GREEN)
    if not red:
        write_out(args.out, {})
        return 0
    client = _load_client()
    if client is None:
        warn("runner-controls/github_rest.py no esta en el catalogo: sin anotaciones")
        return 0
    try:
        runs = fetch_check_runs(client, args.repo, args.sha)
    except (client.GitHubApiError, TypeError) as exc:
        warn(f"no se pudieron leer los check runs de {args.sha}: {exc}")
        return 0
    annotations: dict[str, str] = {}
    for job in red:
        published = published_name(job, check_names, args.event)
        run = find_red_run(runs, published)
        if run is None:
            warn(f"no hay un check rojo con el nombre publicado '{published}' "
                 f"(job '{job}') entre los check runs del SHA: el job se queda "
                 "sin anotacion")
            continue
        try:
            message = first_annotation(client, args.repo, run)
        except (client.GitHubApiError, TypeError) as exc:
            warn(f"no se pudieron leer las anotaciones de '{published}': {exc}")
            continue
        if message:
            annotations[job] = message
    try:
        write_out(args.out, annotations)
    except OSError as exc:
        warn(f"no se pudo escribir el fichero de anotaciones: {exc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
