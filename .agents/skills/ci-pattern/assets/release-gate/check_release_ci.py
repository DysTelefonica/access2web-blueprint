#!/usr/bin/env python3
# ci-pattern asset — release gate: CI en verde sobre el commit que se etiqueta (#1015)
"""Gate de CI que acompaña al corte de release (HR-19).

El gate del tag (``check_release_tag.py``) mide **el tag**: semver estricto,
anotado, versión creciente y commit alcanzable desde la rama por defecto. No
mide el CI. El workflow de release y el de CI son independientes: los dos se
disparan sobre el mismo commit y ninguno espera al otro, así que un tag puede
publicar con el CI en rojo sobre los bytes que se firman.

Este gate mide esa condición y se ejecuta en el job previo del workflow de
release, justo antes de publicar:

- Resuelve localmente, con ``git``, el commit del tag (mismo criterio que el
  gate del tag: ``refs/tags/<tag>^{}``).
- Pregunta al host los runs del workflow de CI (``--workflow``, por defecto
  ``ci.yml``) cuyo ``head_sha`` es ese commit.
- Sin ningún run para ese SHA no dice nada: sale 2. Empujar el commit a una
  rama que ejecute el CI, esperar a que termine y relanzar la release desde el
  mismo tag es lo que lo mide.
- Con runs, toma el más nuevo —el orden de la API no es un contrato: desempata
  por fecha de creación y, a igualdad, por el ID inmutable del run— y exige
  ``status == completed`` y ``conclusion == success``.

Un token ausente, una respuesta que no sea 2xx, un cuerpo que no sea JSON o un
listado que no termina son todos salida 2: el gate no informa «verde» sobre un
sujeto que no ha podido medir (HR-3).

Códigos de salida: 0 el CI del commit del tag terminó en ``success`` · 1 el
último run no terminó en ``success`` (rojo o todavía en curso) · 2 sujeto no
medible (sin runs para ese SHA, API caída, tag inexistente, repositorio que no
es un repositorio git).
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

# El cliente REST vive en `assets/runner-controls/` y viaja con la skill; este
# gate lo reutiliza en vez de abrir su propia vía de red al host.
_CLIENT_DIR = Path(__file__).resolve().parent.parent / "runner-controls"
if str(_CLIENT_DIR) not in sys.path:
    sys.path.insert(0, str(_CLIENT_DIR))

from github_rest import GitHubApiError, call  # noqa: E402

PER_PAGE = 100
MAX_PAGES = 50
EXIT_OK = 0
EXIT_FINDINGS = 1
EXIT_UNMEASURABLE = 2


class GateError(Exception):
    """Una condición que el gate se niega a interpretar; maps a exit 2."""


class ApiUnavailableError(GateError):
    """El host no devolvió una respuesta legible; maps a exit 2."""


def git(repo, *args):
    # git del sistema por nombre y argumentos construidos aquí: no hay shell y
    # no hay entrada que no venga de las opciones de la línea de órdenes.
    argv = ["git", "-C", str(repo), *args]  # noqa: S607
    return subprocess.run(argv, capture_output=True, text=True)  # noqa: S603


def must_git(repo, *args) -> str:
    r = git(repo, *args)
    if r.returncode != 0:
        raise GateError(f"git {' '.join(args)} failed: {r.stderr.strip()}")  # noqa: TRY003
    return r.stdout.strip()


def _pin_output_encoding() -> None:
    """Fija stdout/stderr a UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def tag_commit(repo: str, tag: str) -> str:
    """Commit al que apunta el tag; el mismo criterio que el gate del tag."""
    if git(repo, "rev-parse", "--git-dir").returncode != 0:
        raise GateError(f"{repo!r} is not a git repository (empty subject)")  # noqa: TRY003
    if git(repo, "rev-parse", "-q", "--verify", f"refs/tags/{tag}").returncode != 0:
        raise GateError(f"tag {tag!r} does not exist in the repository (empty subject)")  # noqa: TRY003
    return must_git(repo, "rev-parse", f"refs/tags/{tag}^{{}}")


def fetch_runs(repo: str, workflow_ref: str, sha: str) -> list[dict]:
    """Runs del workflow `workflow_ref` para `sha`. Doble en los tests."""
    runs: list[dict] = []
    for page in range(1, MAX_PAGES + 1):
        path = (f"/repos/{repo}/actions/workflows/{workflow_ref}/runs"
                f"?head_sha={sha}&per_page={PER_PAGE}&page={page}")
        try:
            text, _link = call("GET", path)
            chunk = json.loads(text)
        except GitHubApiError as exc:
            raise ApiUnavailableError(str(exc)) from exc
        except json.JSONDecodeError as exc:
            raise ApiUnavailableError(f"response to {path} is not valid JSON: {exc}") from exc  # noqa: TRY003
        if not isinstance(chunk, dict) or not isinstance(chunk.get("workflow_runs"), list):
            raise ApiUnavailableError(f"response to {path} is not a workflow runs listing")  # noqa: TRY003
        batch = chunk["workflow_runs"]
        runs += [r for r in batch if isinstance(r, dict)]
        if len(batch) < PER_PAGE:
            return runs
    raise ApiUnavailableError(f"listing of runs for {sha} exceeds {MAX_PAGES} pages; "  # noqa: TRY003
                               "refusing to guess which one is the newest")


def newest_run(runs: list[dict]) -> dict:
    """El run más nuevo: la API no ordena, así que se decide con dos datos."""
    return max(runs, key=lambda r: (str(r.get("created_at") or ""), int(r.get("id") or 0)))


def evaluate(runs: list[dict], workflow_ref: str, tag: str) -> tuple[int, str]:
    """Veredicto sobre la lista de runs del SHA del tag."""
    if not runs:
        return EXIT_UNMEASURABLE, (
            f"RELEASE-CI ERROR: no hay ningún run de {workflow_ref} para el commit del "
            f"tag {tag}. Empuje el commit a una rama que ejecute el workflow de CI, "
            "espere a que termine y relance la release desde el mismo tag")
    run = newest_run(runs)
    url = run.get("html_url") or "(sin URL)"
    status = run.get("status") or "none"
    conclusion = run.get("conclusion") or "none"
    if status != "completed":
        return EXIT_FINDINGS, (
            f"RELEASE-CI FAIL: el último run de {workflow_ref} (id {run.get('id')}, "
            f"{url}) está en status {status}, no completado. Espere a que termine y "
            "relance la release desde el mismo tag")
    if conclusion != "success":
        return EXIT_FINDINGS, (
            f"RELEASE-CI FAIL: el último run de {workflow_ref} (id {run.get('id')}, "
            f"{url}) terminó en {conclusion}. Arregle en la rama por defecto, deje el "
            "CI en verde sobre el commit exacto que va a publicar y mueva el tag")
    return EXIT_OK, (f"RELEASE-CI OK: el último run de {workflow_ref} "
                     f"(id {run.get('id')}, {url}) terminó en success sobre el commit del tag")


def main(argv, fetch_runs=None) -> int:
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Release CI gate (HR-19).")
    ap.add_argument("--repo", required=True)
    ap.add_argument("--tag", required=True)
    ap.add_argument("--workflow", default="ci.yml",
                    help="basename del workflow de CI del consumer (por defecto ci.yml)")
    args = ap.parse_args(argv)
    client = fetch_runs or globals()["fetch_runs"]
    try:
        sha = tag_commit(args.repo, args.tag)
        runs = client(args.repo, args.workflow, sha)
        code, message = evaluate(runs, args.workflow, args.tag)
    except GateError as exc:
        code, message = EXIT_UNMEASURABLE, f"RELEASE-CI ERROR: {exc}"
    stream = sys.stderr if code == EXIT_UNMEASURABLE else sys.stdout
    print(message, file=stream)
    return code


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
