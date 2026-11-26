#!/usr/bin/env python3
"""Carga y validación de la instantánea nocturna (#1627, slice 4).

Uso: snapshot.py --snapshot <json>. Sale 0 si el contrato se cumple; 2 si no, nombrando
la clave o el campo, sin traceback. Lo usa el informe nocturno antes de renderizar.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

REQUIRED = ("repository", "night", "last_good_night", "prs", "jobs", "deploy")
PR_KEYS = ("number", "title", "merged_at", "merge_sha", "issues")
JOB_KEYS = ("name", "expensive", "conclusion", "run_url")


class SnapshotError(ValueError):
    """La instantánea no cumple el contrato; el mensaje nombra la clave o el campo."""


class SnapshotShapeError(SnapshotError):
    def __init__(self) -> None:
        super().__init__("la instantánea debe ser un objeto JSON")


class MissingKeyError(SnapshotError):
    def __init__(self, where: str, key: str) -> None:
        super().__init__(f"falta la clave '{key}' en {where}")


class BadDateError(SnapshotError):
    def __init__(self, where: str, key: str, value: object, reason: str) -> None:
        super().__init__(f"'{key}' en {where} no es una fecha válida ({reason}): {value!r}")


class UnreadableSnapshotError(SnapshotError):
    def __init__(self, name: str, reason: str) -> None:
        super().__init__(f"no se puede leer la instantánea {name}: {reason}")


class BadJsonError(SnapshotError):
    def __init__(self, name: str, reason: str) -> None:
        super().__init__(f"JSON mal formado en {name} ({reason})")


def _need(obj: object, key: str, where: str) -> object:
    if not isinstance(obj, dict) or key not in obj:
        raise MissingKeyError(where, key)
    return obj[key]


def when(value: object, key: str, where: str) -> datetime:
    """Una fecha ISO 8601 con zona horaria. Sin zona no se compara con las demás."""
    try:
        moment = datetime.fromisoformat(str(value))
    except ValueError as error:
        raise BadDateError(where, key, value, "formato ISO 8601") from error
    if moment.tzinfo is None:
        raise BadDateError(where, key, value, "falta la zona horaria")
    return moment


def validate(snap: object) -> dict:
    if not isinstance(snap, dict):
        raise SnapshotShapeError()
    for key in REQUIRED:
        _need(snap, key, "la instantánea")
    when(_need(snap["night"], "started_at", "night"), "started_at", "night")
    good = snap["last_good_night"]
    if good is not None:
        when(_need(good, "finished_at", "last_good_night"), "finished_at", "last_good_night")
    _need(snap["deploy"], "deployed", "deploy")
    for job in snap["jobs"]:
        name = job.get("name", "?") if isinstance(job, dict) else "?"
        for key in JOB_KEYS:
            _need(job, key, f"el job {name}")
    for pr in snap["prs"]:
        number = pr.get("number", "?") if isinstance(pr, dict) else "?"
        for key in PR_KEYS:
            _need(pr, key, f"el PR {number}")
        when(pr["merged_at"], "merged_at", f"el PR {number}")
    return snap


def load(path: Path) -> dict:
    """Lee y valida la instantánea; lanza SnapshotError si no cumple el contrato."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        raise UnreadableSnapshotError(path.name, error.strerror or "error de lectura") from error
    try:
        return validate(json.loads(text))
    except json.JSONDecodeError as error:
        raise BadJsonError(path.name, error.msg) from error


def _pin_output_encoding() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--snapshot", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        load(args.snapshot)
    except SnapshotError as error:
        print(f"instantánea no válida: {error}", file=sys.stderr)
        return 2
    print("instantánea válida")
    return 0


if __name__ == "__main__":
    sys.exit(main())
