#!/usr/bin/env python3
# check_checks_table.py — las tablas de checks escritas a mano contra los jobs reales (HR-82)
"""Mide las tablas de checks que un consumer DECLARA contra los workflows reales.

La fase 5 ya mide la tabla del documento operativo GENERADO; las tablas escritas a
mano no se medían, y una tabla escrita a mano es tan vinculante como la generada: el
caso vivo es `DysTelefonica/cadete`, cuya tabla atribuía `php-lint` a `ci.yml` (el
job vive en `php-lint-push.yml`), omitía `ci-images` (que sí es un job requerido) y
una auditoría le añadió un job que no define ningún workflow.

Tres reds distintos, porque el remedio es distinto en cada caso: un job que no
define ningún workflow (una receta fabricada), un job que vive en otro workflow (una
atribución que ya no dice la verdad) y un job requerido que no aparece en ninguna
tabla declarada (una regla del CI que nadie sabe recuperar).

La capa es de solo lectura: el documento declarado, los workflows del repo y la
política de jobs requeridos. Stdlib solamente, sin red. La página del gate es
`assets/checks-table/README.md`.
"""

from __future__ import annotations

import importlib.util
import json
import re
from functools import lru_cache
from pathlib import Path
from typing import NamedTuple

#: Claves que admite la declaración `checks_tables` del contrato de adopción.
CONFIG_KEYS = ("tables",)
#: Claves que admite cada tabla declarada.
TABLE_KEYS = ("path", "workflow", "section")
#: Cabeceras cuya primera columna es la del job.
JOB_HEADERS = frozenset({"job", "jobs", "job de ci", "check", "checks"})
#: Cabeceras cuya segunda columna es la del nombre del check publicado.
CHECK_HEADERS = frozenset({"check", "checks", "nombre del check", "contexto", "context"})
POLICY_PATH = ".github/required-jobs-policy.json"
WORKFLOWS_DIR = ".github/workflows"
#: Cabecera + separador: el mínimo de filas de un bloque para que tenga cabecera.
HEADER_ROWS = 2

_SEPARATOR = re.compile(r":?-{3,}:?")
_TOKEN = re.compile(r"`([^`]+)`")
_HEADING = re.compile(r"^(#+)\s+(.*\S)\s*$")


class Index(NamedTuple):
    """Lo que los workflows del repo declaran: jobs y nombres de check publicados."""

    jobs: dict          # workflow -> {job id}
    names: dict         # workflow -> {nombre de check publicado}
    by_job: dict        # job id -> {workflow}
    by_name: dict       # nombre de check -> {workflow}
    names_by_job: dict  # job id -> {nombre de check publicado}


class Ctx(NamedTuple):
    """Lo que mide la capa: el repo, sus workflows, su índice y sus jobs requeridos."""

    repo: Path
    docs: list
    index: Index
    required: list
    problems: list


@lru_cache(maxsize=1)
def _reader():
    """El lector estructural de workflows del canon (el del gate de política).

    Se importa EN PROCESO, nunca como subproceso (HR-46): un gate que invoca a
    otro gate como subproceso multiplica los procesos sin medir nada más."""
    path = Path(__file__).resolve().parents[1] / "workflow-policy" / "check_workflow_policy.py"
    spec = importlib.util.spec_from_file_location("check_workflow_policy", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _workflows(repo):
    """(workflows, problemas): los workflows legibles y los que no se pudieron leer."""
    docs = []
    problems = []
    directory = repo / WORKFLOWS_DIR
    if not directory.is_dir():
        return docs, [f"no existe el directorio {WORKFLOWS_DIR}"]
    paths = sorted(list(directory.glob("*.yml")) + list(directory.glob("*.yaml")))
    for path in paths:
        rel = path.relative_to(repo).as_posix()
        try:
            doc = _reader().parse_workflow(path.read_text(encoding="utf-8"), rel)
        except (OSError, UnicodeDecodeError, ValueError) as exc:
            problems.append(f"no se pudo leer el workflow {rel}: {exc}")
            continue
        docs.append((rel, doc))
    return docs, problems


def _index(docs):
    """El índice de jobs y de nombres de check que publican los workflows."""
    published = _reader().published_names_by_job(docs)
    jobs, names, by_job, by_name, names_by_job = {}, {}, {}, {}, {}
    for (path, jid), published_names in published.items():
        jobs.setdefault(path, set()).add(jid)
        by_job.setdefault(jid, set()).add(path)
        for name in published_names:
            names.setdefault(path, set()).add(name)
            by_name.setdefault(name, set()).add(path)
            names_by_job.setdefault(jid, set()).add(name)
    return Index(jobs, names, by_job, by_name, names_by_job)


def _required_jobs(repo):
    """(jobs requeridos, problema): la política del repo, o por qué no se pudo leer."""
    path = repo / POLICY_PATH
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return [], f"no se pudo leer {POLICY_PATH}: {exc}"
    jobs = data.get("required_jobs") if isinstance(data, dict) else None
    if not isinstance(jobs, list):
        return [], f"{POLICY_PATH} no declara 'required_jobs' (lista)"
    return [job for job in jobs if isinstance(job, str)], None


def _cells(line):
    """Las celdas de una fila de tabla markdown."""
    inner = line.strip()
    if inner.startswith("|"):
        inner = inner[1:]
    if inner.endswith("|"):
        inner = inner[:-1]
    return [cell.strip() for cell in inner.split("|")]


def _token(cell):
    """El texto entre backticks de una celda, o None si no lo hay."""
    match = _TOKEN.fullmatch(cell.strip())
    return match.group(1).strip() if match else None


def _header(cell):
    """La cabecera de una columna, normalizada para compararla."""
    text = _token(cell)
    if text is None:
        text = re.sub(r"[*_]", "", cell).strip()
    return text.lower()


def _blocks(text):
    """Los bloques de tabla markdown del texto (líneas que empiezan por «|»)."""
    blocks, current = [], []
    for line in text.splitlines():
        if line.lstrip().startswith("|"):
            current.append(line)
        elif current:
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    return blocks


def _block_rows(block):
    """Las filas de jobs de un bloque de tabla; [] si no es una tabla de jobs."""
    cells = [_cells(line) for line in block]
    header = None
    if len(cells) >= HEADER_ROWS and all(_SEPARATOR.fullmatch(cell) for cell in cells[1]):
        header = cells[0]
        cells = cells[2:]
    if header is None:
        check_col = None
    elif _header(header[0]) not in JOB_HEADERS:
        return []
    else:
        check_col = 1 if len(header) > 1 and _header(header[1]) in CHECK_HEADERS else None
    rows = []
    for row in cells:
        job = _token(row[0]) if row else None
        if not job:
            continue
        check = _token(row[check_col]) if check_col is not None and len(row) > check_col else None
        rows.append((job, check))
    return rows


def _rows(text):
    """(job, check) por fila de las tablas de jobs del texto."""
    rows = []
    for block in _blocks(text):
        rows += _block_rows(block)
    return rows


def _section(text, heading):
    """El cuerpo de la sección con ese título; None si no existe."""
    lines = text.splitlines()
    start = None
    level = 0
    for index, line in enumerate(lines):
        match = _HEADING.match(line)
        if match and match.group(2).strip() == heading.strip():
            start = index + 1
            level = len(match.group(1))
            break
    if start is None:
        return None
    end = len(lines)
    for index in range(start, len(lines)):
        match = _HEADING.match(lines[index])
        if match and len(match.group(1)) <= level:
            end = index
            break
    return "\n".join(lines[start:end])


def _doc_text(repo, rel, section):
    """(texto, problema): el texto de la tabla declarada, o por qué no se leyó."""
    path = repo / rel
    if not path.is_file():
        return None, f"la tabla declarada no existe: {rel}"
    text = path.read_text(encoding="utf-8")
    if not section:
        return text, None
    body = _section(text, section)
    if body is None:
        return None, f"la sección «{section}» no existe en {rel}"
    return body, None


def _resolve_workflow(declared, docs):
    """El workflow declarado (ruta o fichero), o None si no existe o es ambiguo."""
    paths = [path for path, _doc in docs]
    if declared in paths:
        return declared
    matches = [path for path in paths if Path(path).name == declared]
    return matches[0] if len(matches) == 1 else None


def _row_findings(rel, workflow, job, check, index):
    """Una fila contra el workflow declarado: está, se movió, o no existe."""
    if job in index.jobs.get(workflow, set()):
        return []
    if check and check in index.names.get(workflow, set()):
        return []
    elsewhere = set(index.by_job.get(job, set()))
    if check:
        elsewhere |= index.by_name.get(check, set())
    elsewhere.discard(workflow)
    if elsewhere:
        donde = Path(sorted(elsewhere)[0]).name
        return [f"la tabla de '{rel}' atribuye el job '{job}' al workflow "
                f"'{Path(workflow).name}', pero lo define '{donde}': el job se movió "
                "de workflow y la tabla ya no dice la verdad"]
    return [f"la tabla de '{rel}' documenta el job '{job}', que no define ningún "
            f"workflow de {WORKFLOWS_DIR}: una receta fabricada es peor que la ausencia"]


def _table_findings(ctx, entry):
    """(hallazgos, filas) de una tabla declarada."""
    if not isinstance(entry, dict):
        return ["cada tabla declarada debe ser un objeto {path, workflow}"], []
    findings = [f"clave desconocida en la tabla declarada: {key}"
                for key in sorted(set(entry) - set(TABLE_KEYS))]
    path = entry.get("path")
    workflow = entry.get("workflow")
    if not isinstance(path, str) or not path.strip() \
            or not isinstance(workflow, str) or not workflow.strip():
        return findings + ["cada tabla declarada necesita 'path' y 'workflow'"], []
    rel = path.strip()
    key = _resolve_workflow(workflow.strip(), ctx.docs)
    if key is None:
        return findings + [f"el workflow '{workflow.strip()}' declarado para '{rel}' no "
                           f"existe en {WORKFLOWS_DIR}"], []
    text, problem = _doc_text(ctx.repo, rel, entry.get("section"))
    if problem:
        return findings + [problem], []
    rows = _rows(text)
    if not rows:
        return findings + [f"la tabla declarada en '{rel}' no tiene filas de jobs: "
                           "una tabla sin jobs no mide nada"], []
    row_findings = []
    for job, check in rows:
        row_findings += _row_findings(rel, key, job, check, ctx.index)
    return findings + row_findings, rows


def _coverage_findings(rows, required, index):
    """Todo job requerido tiene que aparecer en alguna tabla declarada."""
    documented = {job for job, _check in rows}
    findings = []
    for job in required:
        names = index.names_by_job.get(job, set())
        covered = job in documented or any(check and check in names for _job, check in rows)
        if not covered:
            findings.append(f"el job requerido '{job}' no aparece en ninguna tabla de "
                            "checks declarada: ninguna regla del CI puede fallar sin "
                            "estar documentada")
    return findings


def check(repo, decl):
    """Los hallazgos de la capa: las tablas declaradas contra los jobs reales."""
    tables = decl.get("tables")
    if not isinstance(tables, list) or not tables:
        return ["falta 'tables' (lista no vacía de {path, workflow}): sin ella no hay "
                "tabla que medir"]
    docs, problems = _workflows(repo)
    required, policy_problem = _required_jobs(repo)
    if policy_problem:
        problems.append(policy_problem)
    ctx = Ctx(Path(repo), docs, _index(docs), required, problems)
    findings = list(ctx.problems)
    rows_all = []
    for entry in tables:
        table_findings, rows = _table_findings(ctx, entry)
        findings += table_findings
        rows_all += rows
    if docs:
        findings += _coverage_findings(rows_all, ctx.required, ctx.index)
    return findings
