#!/usr/bin/env python3
# ci-pattern asset — comprobación diaria de deploy a producción (#1627 slice 7)
"""Comprueba que hoy (día UTC) hubo al menos un deploy exitoso al entorno.

Un deploy cuenta si es un GitHub Deployment del entorno (por defecto
`production`, `--environment`) creado ese día UTC y cuyo estado más reciente
es `success`. Si ninguno cumple, abre UNA issue de máxima prioridad con la
etiqueta fija `prerelease:night-red` (`--label` la sustituye). Si ya hay una
issue abierta con esa etiqueta y ese mismo título, no abre otra.

`--date` debe ser una fecha real (YYYY-MM-DD). `--dry-run` reporta lo que haría
y no crea nada. La API se llama por HTTP con el cliente compartido
`runner-controls/github_rest.py` (`GITHUB_TOKEN`/`GITHUB_API_URL` del entorno):
el runner del consumer no trae el CLI `gh` (#1691).

Códigos de salida: 0 hubo deploy exitoso; 1 no lo hubo (issue abierta, ya
abierta o en ensayo); 2 la API no se pudo leer, devolvió una forma inesperada o
la fecha es inválida (fallo cerrado, sin issue).
"""
import argparse
import importlib.util
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

DEFAULT_LABEL = "prerelease:night-red"
DEFAULT_ENVIRONMENT = "production"
FAIL_CLOSED = 2
LIST_LIMIT = 200  # issues revisadas en total, como el `--limit 200` de gh
PER_PAGE = 100  # máximo por página de la API de GitHub


def _load_github_rest():
    """Carga el cliente compartido `runner-controls/github_rest.py` por ruta (#1691).

    El asset vive en `assets/prerelease-deploy/` y el cliente en
    `assets/runner-controls/`: se carga por ruta, sin depender del PYTHONPATH del
    runner. Sin él, el script no arranca: sale con 2 (HR-3).
    """
    rest_path = Path(__file__).resolve().parents[1] / "runner-controls" / "github_rest.py"
    if not rest_path.is_file():
        return None
    spec = importlib.util.spec_from_file_location("prerelease_deploy_github_rest", str(rest_path))
    if spec is None or spec.loader is None:
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # noqa: BLE001 - un cliente roto no puede romper el arranque
        print(f"prerelease-deploy: runner-controls/github_rest.py no se pudo cargar: {exc}",
              file=sys.stderr)
        sys.exit(FAIL_CLOSED)
    return module


_REST = _load_github_rest()
if _REST is None:
    print("prerelease-deploy: runner-controls/github_rest.py no esta en el catalogo "
          "(HR-3 fail-closed)", file=sys.stderr)
    sys.exit(2)


class ApiError(Exception):
    """La API no se pudo leer o devolvió una forma inesperada (fallo cerrado)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def _read_json(path, what):
    """GET de un listado por la API; falla cerrado ante error o forma inesperada."""
    try:
        data = _REST.get_json(path)
    except _REST.GitHubApiError as exc:
        reason = f"la API no devuelve {what}: {exc}"
        raise ApiError(reason) from exc
    if not isinstance(data, list) or not all(isinstance(item, dict) for item in data):
        reason = f"{what} no tiene la forma esperada"
        raise ApiError(reason)
    return data


def _read_issues(repo, label):
    """Issues abiertas con la etiqueta, hasta LIST_LIMIT: la API pagina por `page`."""
    items = []
    for page in range(1, LIST_LIMIT // PER_PAGE + 1):
        chunk = _read_json(
            f"repos/{repo}/issues?labels={quote(label, safe='')}&state=open"
            f"&per_page={PER_PAGE}&page={page}", "las issues abiertas")
        items += chunk
        if len(chunk) < PER_PAGE:
            break
    return items


def parse_date(text):
    """Una fecha real YYYY-MM-DD; 2026-13-45 no lo es."""
    try:
        datetime.strptime(text, "%Y-%m-%d")
    except ValueError as exc:
        reason = f"--date debe ser una fecha real YYYY-MM-DD, no {text!r}"
        raise argparse.ArgumentTypeError(reason) from exc
    return text


def parse(argv):
    p = argparse.ArgumentParser(description="Comprobación diaria de deploy a producción (#1627 s7).")
    p.add_argument("--repo", required=True, help="owner/name del repositorio")
    p.add_argument("--environment", default=DEFAULT_ENVIRONMENT, help="entorno del deploy (por defecto production)")
    p.add_argument("--date", type=parse_date, default=None, help="día UTC YYYY-MM-DD (por defecto hoy en UTC)")
    p.add_argument("--label", default=DEFAULT_LABEL, help="etiqueta fija de la issue")
    p.add_argument("--dry-run", action="store_true", help="reporta sin crear nada")
    return p.parse_args(argv)


def successful_deploy_on(repo, environment, day):
    """El id del deployment del día con estado más reciente success, o None."""
    deployments = _read_json(
        f"repos/{repo}/deployments?environment={quote(environment, safe='')}&per_page={PER_PAGE}",
        "los deployments")
    for dep in deployments:
        if dep.get("environment") != environment or str(dep.get("created_at", ""))[:10] != day:
            continue
        statuses = _read_json(
            f"repos/{repo}/deployments/{dep.get('id')}/statuses?per_page=1",
            f"los estados del deployment {dep.get('id')}")
        if statuses and statuses[0].get("state") == "success":
            return dep.get("id")
    return None


def open_issue_or_create(a, day):
    title = f"Sin deploy a {a.environment} exitoso el {day}"
    listed = _read_issues(a.repo, a.label)
    if any(item.get("title") == title for item in listed):
        return 1, f"ya abierta: existe una issue abierta con la etiqueta {a.label} y el título «{title}»"
    body = (
        f"Prioridad: máxima. No hubo ningún deploy exitoso a {a.environment} el {day} (UTC).\n\n"
        "Las issues con esta etiqueta van antes que el trabajo nuevo. Se vuelve a lanzar la\n"
        "canalización nocturna el mismo día, y así hasta que el deploy salga en verde."
    )
    if a.dry_run:
        return 1, (f"dry-run: POST repos/{a.repo}/issues con title {title!r} y label {a.label!r}; "
                   "no se crea nada")
    try:
        # Un POST que crea una issue no se reintenta: repetirlo duplicaría la issue (#1592).
        text, _ = _REST.call("POST", f"repos/{a.repo}/issues",
                             {"title": title, "body": body, "labels": [a.label]})
    except _REST.GitHubApiError as exc:
        return FAIL_CLOSED, f"fail-closed: la API no crea la issue: {exc}"
    return 1, _created_url(text)


def _created_url(text):
    """La URL de la issue creada; si la API no la devuelve, el texto crudo."""
    try:
        data = json.loads(text or "{}")
    except ValueError:
        return (text or "").strip()
    if isinstance(data, dict) and data.get("html_url"):
        return str(data["html_url"])
    return (text or "").strip()


def decide(a, day):
    """(codigo, mensaje): 0 si hubo deploy hoy; si no, abre o confirma la issue."""
    try:
        dep_id = successful_deploy_on(a.repo, a.environment, day)
        if dep_id is not None:
            return 0, f"ok: deploy a {a.environment} exitoso el {day} (deployment {dep_id})"
        return open_issue_or_create(a, day)
    except ApiError as exc:
        return FAIL_CLOSED, f"fail-closed: {exc}"


def main(argv=None):
    _pin_output_encoding()
    a = parse(argv if argv is not None else sys.argv[1:])
    day = a.date or time.strftime("%Y-%m-%d", time.gmtime())
    code, msg = decide(a, day)
    print(msg, file=sys.stderr if code == FAIL_CLOSED else sys.stdout)
    return code


if __name__ == "__main__":
    sys.exit(main())
