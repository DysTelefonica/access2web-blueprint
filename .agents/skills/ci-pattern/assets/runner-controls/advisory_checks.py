#!/usr/bin/env python3
# ci-pattern asset -- los checks advisory declarados por el consumer (DysTelefonica/team-skills#1762)
"""Los checks que el consumer declara como no bloqueantes (`advisory_checks`).

La clave es una lista de patrones de nombre (globs, `fnmatch`) en
`.github/host-contract.json`. Los check runs que casan no cuentan para el
veredicto de los guards de merge: un job espejo puede salir en rojo, o con
`startup_failure`, sin parar los merges de los agentes (APAP_WEB#4411, #4423).
Sin la clave no hay advisory; un check requerido nunca queda tapado por un
patron; una declaracion invalida falla cerrado (HR-3).
"""
from __future__ import annotations

import fnmatch
import json
from pathlib import Path
from typing import NamedTuple

CONTRACT_PATH = ".github/host-contract.json"


class AdvisoryContractError(Exception):
    """La declaracion de checks advisory no tiene una forma valida (falla cerrado)."""

    template = "the advisory_checks declaration is not valid: {0}"

    def __init__(self, *details: object) -> None:
        super().__init__(self.template.format(*details))


class ContractNotAnObjectError(AdvisoryContractError):
    template = "the host contract is not a JSON object"


class ContractUnreadableError(AdvisoryContractError):
    template = "could not read the host contract {0}: {1}"


class ContractInvalidJsonError(AdvisoryContractError):
    template = "the host contract {0} is not valid JSON: {1}"


class AdvisoryShapeError(AdvisoryContractError):
    template = "advisory_checks must be a list of non-empty glob patterns (got {0!r})"


class RequiredChecksNotListError(AdvisoryContractError):
    # #1795: misma regla que check_push_compliance.required_checks().
    template = "required_checks must be a list (got {0!r})"


class RequiredCheckEntryError(AdvisoryContractError):
    # #1795: una entrada rota se nombra, no se descarta: descartarla dejaria sin
    # proteger ese requerido frente a un patron advisory.
    template = "required_checks[{0}] must be an object with a non-empty string 'name' (got {1!r})"


def advisory_patterns(contract: object) -> tuple[str, ...]:
    """Los patrones `advisory_checks` de un host-contract ya leido."""
    if not isinstance(contract, dict):
        raise ContractNotAnObjectError()
    if "advisory_checks" not in contract:
        return ()
    value = contract["advisory_checks"]
    if not isinstance(value, list) or not all(isinstance(p, str) and p.strip() for p in value):
        raise AdvisoryShapeError(value)
    return tuple(p.strip() for p in value)


def required_check_names(contract: object) -> tuple[str, ...]:
    """Los checks que el contrato declara como requeridos (`required_checks`).

    Un check requerido nunca queda tapado por un patron advisory, venga el
    requerimiento de la linea de comandos o del contrato. Sin la clave no hay
    requeridos (#1795). Una forma invalida falla cerrado y nombra la entrada rota:
    la misma regla que `check_push_compliance.required_checks()`."""
    if not isinstance(contract, dict) or "required_checks" not in contract:
        return ()
    raw = contract["required_checks"]
    if not isinstance(raw, list):
        raise RequiredChecksNotListError(raw)
    names = []
    for index, entry in enumerate(raw):
        name = entry.get("name") if isinstance(entry, dict) else None
        # Un nombre solo con espacios no puede casar con ningun check run: es una entrada rota.
        if not isinstance(name, str) or not name.strip():
            raise RequiredCheckEntryError(index, entry)
        names.append(name)
    return tuple(names)


class ContractAdvisories(NamedTuple):
    """Lo que el host-contract declara: los checks advisory y los requeridos."""

    patterns: tuple[str, ...]
    required: tuple[str, ...]


def read_contract(path: object = CONTRACT_PATH) -> ContractAdvisories:
    """El host-contract del consumer: patrones advisory y checks requeridos.

    Si el fichero no existe, no hay declaracion (el comportamiento es el de siempre).
    Si existe pero no se puede leer (no es UTF-8, su JSON esta roto, o su declaracion
    no tiene una forma valida), falla cerrado (HR-3)."""
    target = Path(path)
    try:
        text = target.read_text(encoding="utf-8")
    except FileNotFoundError:
        return ContractAdvisories((), ())
    except (OSError, UnicodeDecodeError) as exc:
        raise ContractUnreadableError(target, exc) from exc
    try:
        contract = json.loads(text)
    except ValueError as exc:
        raise ContractInvalidJsonError(target, exc) from exc
    return ContractAdvisories(advisory_patterns(contract), required_check_names(contract))


def read_advisory_patterns(path: object) -> tuple[str, ...]:
    """Los patrones `advisory_checks` del host-contract de `path` (si no existe, ninguno)."""
    return read_contract(path).patterns


def is_advisory(name: str, patterns: tuple[str, ...]) -> bool:
    """Whether a check run name matches an advisory pattern."""
    return any(fnmatch.fnmatchcase(name, pattern) for pattern in patterns)


def split_runs(runs: list, patterns: tuple[str, ...], required=()) -> tuple[list, list]:
    """(bloqueantes, advisory): los runs que se juzgan y los que solo se listan.

    Un nombre requerido nunca es advisory: sigue bloqueando aunque case con un patron."""
    required = set(required)
    blocking, advisory = [], []
    for run in runs:
        name = run.get("name") if isinstance(run, dict) else None
        if isinstance(name, str) and name not in required and is_advisory(name, patterns):
            advisory.append(run)
        else:
            blocking.append(run)
    return blocking, advisory
