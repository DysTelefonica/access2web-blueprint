#!/usr/bin/env python3
# ci-pattern asset — governed merge (DysTelefonica/team-skills#303)
"""Merge gobernado: el unico camino de merge sobre una rama protegida.

El workflow que corre en el runner propio invoca este asset con el numero
de PR; verifica el contrato antes de fusionar: los ``required_checks`` del
contrato en verde sobre la cabeza del PR y ``Closes #N`` solo en el punta de
cadena (un PR
``chain:partial`` lo lleva en la violacion). Un check requerido que aun no
existe o no ha terminado se espera hasta ``--wait-seconds`` (#787). El tamano
no se repite aqui: lo mide ``pr-size`` (excluye lo vendorizado), y ``required``
ya lo incluye; ``--max-lines`` se acepta e ignora por compatibilidad. Si algo falta no fusiona y
devuelve 1 con el hallazgo; el entorno insalvable falla cerrado con 2
(HR-3). Con todo en verde fusiona por la API con la identidad del bot que
ejecuta el workflow, cierra las issues enlazadas por palabras de cierre y devuelve 0.
El merge con ``GITHUB_TOKEN`` no dispara el cierre implicito de GitHub (0 de 13, #800):
el cierre es explicito, exige ``issues: write`` y, si falla con el merge ya hecho,
devuelve 3 («fusionado, issue sin cerrar»), nunca el 2. Un PR atrasado respecto de su base se bloquea antes (1): el arbol a fusionar
no seria el que reviso el CI (#787) y, si ambos lados cambiaron el mismo fichero
de ``.github/workflows/``, ``GITHUB_TOKEN`` no tiene el permiso ``workflows`` (#783).
Tras el merge, el job mide el arbol integrado y publica el veredicto (#1265):
compara el arbol del commit de merge con el del head verificado (leidos de
``commit.tree.sha``, con el lector compartido del detective post-push) y lo
escribe en stdout y en ``$GITHUB_STEP_SUMMARY``. Arboles iguales: 0. Arboles
distintos: 5 (el merge no se deshace; el rojo dice que el arbol integrado no es
el que reviso el CI). Si la medicion no se puede hacer, falla cerrado con 2
nombrando la causa: un commit ilegible nunca se reporta como coincidencia.
Esa medicion no interrumpe el trabajo posterior al merge (#1279): si falla, el
resumen dice ``not measured: <causa>``, el cierre de las issues enlazadas y el
dispatch post-merge se ejecutan igual, y el 2 sale al final con la causa.
La operacion completa, con los codigos de salida y que hacer con el rojo, esta
en ``references/integrated-tree-evidence.md`` (raiz del skill).
Runner-enforced (HR-34); lo que no pase
por aqui lo detecta el detective post-push (HR-49). Solo stdlib; cada
llamada a ``gh api`` lleva timeout.
"""
import argparse
import base64
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import NamedTuple

from check_push_compliance import (
    SHA_RE,
    FailClosedError,
    check_verdict,
    commit_tree_sha,
    fail,
    gh_json,
    load_json,
    protected_branches,
    required_checks,
)
from github_rest import MAX_PAGES, GitHubApiError, call, get_json

CLOSING = re.compile(r"\b(closes|fixes|resolves)\s+#?\d+", re.IGNORECASE)
WORKFLOWS_DIR = ".github/workflows/"
FORBIDDEN = 403
RACE_CLOSED = 422  # PATCH devolvio 422 = carrera con el cierre auto por `Closes` (#835)
WAITABLE = (": missing", ": pending")  # el CI aun no publico o no termino; un fallo no se espera
DEFAULT_WAIT, DEFAULT_POLL = 300.0, 10.0
POST_DISPATCH_FAILED = 4  # merged, post-dispatch failed (no rollback, #817)
# #1265: merged, the integrated tree is NOT the tree CI reviewed (no rollback).
INTEGRATED_TREE_DRIFT = 5
# #1265: el texto del veredicto. Va al resumen del job y al log; el drift se
# dice con un NOT explicito para que nadie lo lea como un match.
VERDICT_MATCH = "the integrated tree is the tree CI reviewed"
VERDICT_DRIFT = "the integrated tree is NOT the tree CI reviewed"
# #1279: como se dice que la medicion no se pudo hacer. Va al resumen del job
# con la causa al lado: el resumen no afirma una coincidencia que no se midio.
NOT_MEASURED = "not measured"
WORKFLOWS_HINT = ("the token lacks the 'workflows' permission, which GITHUB_TOKEN cannot "
                  "have: GitHub refuses a merge that creates new workflow-file content; "
                  "update the PR branch onto its base or merge it as the maintainer")


def gh_merge(repo: str, number: object, sha: str) -> str:
    # El sha verificado viaja con el merge: si el head cambia entre la
    # comprobación y el PUT, la API responde 409 y el fallo es cerrado (2) —
    # nunca se fusiona una cabeza que nadie verificó (TOCTOU).
    try:
        text, _ = call("PUT", f"repos/{repo}/pulls/{number}/merge",
                       {"commit_title": f"Merge PR #{number} (governed merge)", "sha": sha})
    except GitHubApiError as exc:
        hint = f" ({WORKFLOWS_HINT})" if exc.status == FORBIDDEN and "workflow" in str(exc).lower() else ""
        fail(f"the merge call failed (fail closed): {exc}{hint}", exc)
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        fail(f"the merge API returned no valid JSON body: {exc}", exc)
    if not isinstance(data, dict) or data.get("merged") is not True \
            or not isinstance(data.get("sha"), str):
        fail(f"the merge response does not report a merged PR: {text.strip()!r}")
    return data["sha"]


class MergedNotClosedError(Exception):
    """The merge happened but a linked issue could not be listed or closed (exit 3).

    Each subclass owns its message template; raise sites pass data only.
    """

    template = "merged, linked issue not closed"

    def __init__(self, *details: object) -> None:
        super().__init__(self.template.format(*details))


class IssueListFailedError(MergedNotClosedError):
    template = "could not list the issues linked to PR #{0}: {1}"


class IssuesNotClosedError(MergedNotClosedError):
    template = "linked issue(s) not closed: {0}"


class PostDispatchFailedError(Exception):
    """Al menos un workflow de ``post_merge_dispatch`` respondio 4xx/5xx (exit 4).

    El merge ya esta hecho; no se deshace. El caller termina el job en rojo y
    nombra el workflow que fallo. La superficie de errores coincide con
    ``MergedNotClosedError`` (post-merge, sin rollback) pero con causa distinta:
    este deja el merge intacto y la issue intacta; solo el disparo falla.
    """

    template = "post-merge dispatch failed for {0}: {1}"

    def __init__(self, label: str, reason: str) -> None:
        super().__init__(self.template.format(label, reason))


CLOSE_QUERY = ("query($owner:String!,$name:String!,$pr:Int!){repository(owner:$owner,name:$name)"
               "{pullRequest(number:$pr){closingIssuesReferences(first:50){nodes{number state "
               "repository{nameWithOwner}}}}}}")
ISSUES_HINT = "the workflow needs 'issues: write' to close issues with GITHUB_TOKEN"


def linked_issues(repo: str, number: object) -> list[dict]:
    """Issues que GitHub enlaza al PR por palabras de cierre (su propia interpretacion)."""
    owner, name = repo.split("/")
    variables = {"owner": owner, "name": name, "pr": int(number)}
    try:
        text, _ = call("POST", "graphql", {"query": CLOSE_QUERY, "variables": variables})
        data = json.loads(text)
        nodes = data["data"]["repository"]["pullRequest"]["closingIssuesReferences"]["nodes"]
    except (GitHubApiError, json.JSONDecodeError, KeyError, TypeError) as exc:
        raise IssueListFailedError(number, exc) from exc
    return [node for node in nodes if isinstance(node, dict)]


def default_branch(repo: str, pull: dict) -> str | None:
    named = (pull["base"].get("repo") or {}).get("default_branch")
    if isinstance(named, str):
        return named
    data = gh_json(f"repos/{repo}")
    value = data.get("default_branch") if isinstance(data, dict) else None
    return value if isinstance(value, str) else None


def close_linked_issues(repo: str, pull: dict, number: object, sha: str) -> None:
    """Cierra (completed) las issues abiertas de este repo que el PR enlaza con palabras de cierre.

    Solo si la base es la rama por defecto, como GitHub. Una issue ya cerrada o de otro
    repo se deja; el fallo de cada una se acumula y se informa junto al resto (exit 3).

    #835: cuando ``PATCH /repos/{owner}/{repo}/issues/{n}`` devuelve 422, el
    caso típico no es de permisos (la pista «issues: write» era engañosa) sino
    una carrera con el cierre automático de GitHub por la palabra Closes del
    propio PR mergeado. Se relee la issue y, si ya está cerrada, se trata
    como éxito. Un re-read fallido o que devuelve ``state == "open"`` se
    mantiene como fallo pendiente (exit 3) para no tragarnos defectos reales.
    """
    if pull["base"]["ref"] != default_branch(repo, pull):
        return
    pending = []
    for node in linked_issues(repo, number):
        issue = node.get("number")
        same_repo = ((node.get("repository") or {}).get("nameWithOwner") or "").lower() == repo.lower()
        if not same_repo or node.get("state") != "OPEN" or not isinstance(issue, int):
            continue
        raced_already_closed = False
        try:
            call("PATCH", f"repos/{repo}/issues/{issue}", {"state": "closed", "state_reason": "completed"})
        except GitHubApiError as exc:
            if exc.status == RACE_CLOSED and _issue_now_closed(repo, issue):
                raced_already_closed = True
            else:
                pending.append(f"#{issue} ({exc}; {ISSUES_HINT})")
                continue
        if raced_already_closed:
            print(f"CLOSED: #{issue} (raced; already closed by GitHub's Closes auto-close, #835)")
        else:
            print(f"CLOSED: #{issue} (linked by PR #{number})")
        try:
            call("POST", f"repos/{repo}/issues/{issue}/comments",
                 {"body": f"Closed by PR #{number} (merge {sha})."})
        except GitHubApiError as exc:
            print(f"WARNING: closed #{issue} but could not comment: {exc}", file=sys.stderr)
    if pending:
        raise IssuesNotClosedError(", ".join(pending))


def _issue_now_closed(repo: str, issue: int) -> bool:
    """#835: tras un PATCH 422, re-lee la issue. ``True`` si la API la da cerrada.

    Un error del GET o un cuerpo inválido se interpreta conservadoramente como
    «sigue abierta»: preferimos un exit 3 falso a tragarnos un PATCH que no
    diagnosticamos. El caller ya tiene el 422 original en el log si el re-read
    también falla.
    """
    try:
        text, _ = call("GET", f"repos/{repo}/issues/{issue}", {})
    except GitHubApiError:
        return False
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return False
    return isinstance(data, dict) and data.get("state") == "closed"


# #817: post-merge dispatch (workflow_dispatch rebinds the post-merge run
# set when `GITHUB_TOKEN` suppresses push). The contract may declare
# ``post_merge_dispatch`` as a list of workflow file basenames. Each name
# must exist in the repo and the workflow YAML must declare
# ``workflow_dispatch``; otherwise the dispatch 404s and the run loses
# coverage. Validation runs pre-merge (block); firing runs post-merge
# (any 4xx/5xx is exit 4 with the name, the merge is NOT rolled back).


def _workflow_basename(path: str) -> str:
    """The contract key of a workflow: its file name without extension.

    ``post_merge_dispatch`` in the host-contract names workflows the way the
    catalog does (``deploy``, ``push-compliance``, ``runner-merge``), so the
    listing is normalised to that form to match a contract entry against a
    workflow. This is NOT the identifier the dispatch endpoint accepts: that one
    is the file name with its extension, or the numeric id (#1086).
    """
    base = path.rsplit("/", 1)[-1]
    for suffix in (".yml", ".yaml"):
        if base.endswith(suffix):
            return base[: -len(suffix)]
    return base


def _workflow_dispatch_id(path: str) -> str:
    """The identifier the dispatch endpoint accepts: the workflow FILE NAME (#1086).

    ``POST /repos/{repo}/actions/workflows/{workflow_id}/dispatches`` resolves
    ``workflow_id`` as the workflow file name (``ci.yml``), its path
    (``.github/workflows/ci.yml``) or its numeric id. The basename without the
    extension is not a workflow: the endpoint answers 404, the post-merge
    detective never fires, and the merge stays unaudited (#1086).

    The id is read from the workflows listing, so a ``.yaml`` workflow is
    dispatched as ``deploy.yaml`` and not as a guessed ``.yml``.
    """
    return path.rsplit("/", 1)[-1]


def declares_workflow_dispatch(yaml_text: str) -> bool:
    """Detecta ``workflow_dispatch`` como trigger en un workflow YAML de Actions.

    Suficiente para un gate de validacion previo al merge: escanea el YAML
    linea por linea buscando el bloque ``on:`` a nivel superior (columna 0)
    y confirma que ``workflow_dispatch`` aparece como subclave. No es un
    parser YAML completo — la forma que el catalogo asume (clave ``on:`` a
    columna 0 con sub-claves a dos espacios, con o sin ``push`` /
    ``pull_request`` u otras) es determinista. Si algun workflow del
    consumer adopta flow style (``on: {workflow_dispatch:}``), se reemplaza
    por un parser real sin cambiar el contrato.
    """
    in_on = False
    for line in yaml_text.splitlines():
        if not line or line.startswith("#"):
            continue
        if line.startswith((" ", "\t")):
            if in_on and line.strip().startswith("workflow_dispatch"):
                return True
            continue
        if line.startswith(("on:", "on :")):
            in_on = True
            continue
        in_on = False
    return False


# #933: sangrias de la forma del catalogo: ``on:`` a columna 0, ``workflow_dispatch:`` a
# dos espacios, ``inputs:`` a cuatro.
_DISPATCH_INDENT = 2
_INPUTS_INDENT = 4


def dispatch_input_keys(yaml_text: str) -> set[str]:
    """#933: claves declaradas bajo ``on: workflow_dispatch: inputs:``.

    Mismo escaneo determinista que ``declares_workflow_dispatch``: ``on:`` a columna 0,
    ``workflow_dispatch:`` a dos espacios, ``inputs:`` a cuatro y las claves a la
    indentacion de la primera. Sin ``inputs:`` devuelve un conjunto vacio.
    """
    keys: set[str] = set()
    in_on = in_dispatch = in_inputs = False
    key_indent = None
    for line in yaml_text.splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        indent = len(line) - len(line.lstrip())
        stripped = line.strip()
        if indent == 0:
            in_on = line.startswith(("on:", "on :"))
            in_dispatch = in_inputs = False
            continue
        if not in_on:
            continue
        if indent == _DISPATCH_INDENT:
            in_dispatch, in_inputs = stripped.startswith("workflow_dispatch"), False
            continue
        if not in_dispatch:
            continue
        if indent == _INPUTS_INDENT:
            in_inputs, key_indent = stripped.startswith("inputs:"), None
            continue
        if in_inputs and stripped.endswith(":"):
            key_indent = indent if key_indent is None else key_indent
            if indent == key_indent:
                keys.add(stripped[:-1])
    return keys


def _range_inputs(name: str, yaml_text: str) -> tuple[str | None, bool]:
    """#933: ``before`` y ``after`` van juntos o no van; uno solo deja el rango a medias."""
    range_inputs = dispatch_input_keys(yaml_text) & {"before", "after"}
    if len(range_inputs) == 1:
        only = next(iter(range_inputs))
        return (f"workflow '{name}' declares only the input '{only}' under workflow_dispatch: "
                "the dispatch range needs both 'before' and 'after' (#933)"), False
    return None, range_inputs == {"before", "after"}


def _validate_one_workflow(repo: str, name: str,
                           by_basename: dict[str, str]) -> tuple[str | None, bool]:
    """Valida un nombre declarado en ``post_merge_dispatch``.

    Devuelve ``(hallazgo, con_rango)``. El hallazgo es ``None`` cuando el workflow
    declara ``on: workflow_dispatch:``; en otro caso es el texto que el caller añade
    a la lista de hallazgos. ``con_rango`` indica que el workflow declara ``before`` y
    ``after`` como entradas (#933) y recibe el rango en el dispatch. Las verificaciones
    viven aquí para que el motor de ``validate_post_merge_dispatch`` quede por debajo
    del umbral de C901/PLR0912 del ratchet shrink-only de ruff extendido (#817).
    """
    path = by_basename.get(name)
    if not isinstance(path, str):
        return (f"post_merge_dispatch entry '{name}' is not a workflow in this repo "
                f"(available: {sorted(by_basename)})"), False
    try:
        contents = get_json(f"repos/{repo}/contents/{path}")
    except GitHubApiError as exc:
        return (f"could not read the workflow '{name}' to validate workflow_dispatch: "
                f"{exc}"), False
    if (not isinstance(contents, dict) or contents.get("encoding") != "base64"
            or not isinstance(contents.get("content"), str)):
        return (f"workflow '{name}' returned an unexpected contents shape; the file may "
                "have been removed between the listing and the read"), False
    try:
        yaml_text = base64.b64decode(contents["content"]).decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        return f"workflow '{name}' body is not valid base64/utf-8: {exc}", False
    if not declares_workflow_dispatch(yaml_text):
        return (f"workflow '{name}' does not declare 'workflow_dispatch': the dispatch "
                "step would 404; add 'on: workflow_dispatch:' to its YAML before listing "
                "it in post_merge_dispatch (#817)"), False
    return _range_inputs(name, yaml_text)


WORKFLOWS_PER_PAGE = 100  # el maximo que acepta el listado; el defecto (30) da falsos negativos (#966)
WORKFLOWS_SHAPE = 'an object {"total_count": N, "workflows": [...]}, not a bare list'


class WorkflowsListingError(Exception):
    """El listado de workflows no sirve para validar ``post_merge_dispatch`` (#966).

    El mensaje nombra la forma que espera el gate para que el operador sepa que
    corregir. Lo captura ``validate_post_merge_dispatch`` y se vuelve un hallazgo
    que bloquea el merge (exit 1), igual que un fallo de la llamada.
    """

    template = "the workflows listing could not be read: {0}"

    def __init__(self, *details: object) -> None:
        super().__init__(self.template.format(*details))


class WorkflowsListingShapeError(WorkflowsListingError):
    template = "workflows listing is not the API shape ({0}): expected {1} (#966)"


class WorkflowsListingTooLongError(WorkflowsListingError):
    template = ("the workflows listing of '{0}' exceeds {1} pages ({2} workflows); "
                "refusing to continue")


def list_workflows(repo: str) -> list[object]:
    """Workflows del repo, todas las paginas y con la forma real de la API (#966).

    ``GET repos/{repo}/actions/workflows`` responde un objeto ``{"total_count": N,
    "workflows": [...]}``, no una lista; exigir la lista bloqueaba con
    ``workflows listing is not the API shape: dict`` todo consumer con
    ``post_merge_dispatch`` declarado. Pide ``per_page=100`` y sigue con ``page``
    mientras la pagina venga llena, asi un repo con mas de 30 workflows no da un
    falso «no existe». Sin la clave ``workflows``, o con un valor que no sea
    lista, falla cerrado (HR-3) nombrando la forma esperada.
    """
    listed: list[object] = []
    for page in range(1, MAX_PAGES + 1):
        data = get_json(f"repos/{repo}/actions/workflows?per_page={WORKFLOWS_PER_PAGE}&page={page}")
        items = data.get("workflows") if isinstance(data, dict) else None
        if not isinstance(items, list):
            raise WorkflowsListingShapeError(type(data).__name__, WORKFLOWS_SHAPE)
        listed += items
        if len(items) < WORKFLOWS_PER_PAGE:
            return listed
    raise WorkflowsListingTooLongError(repo, MAX_PAGES, MAX_PAGES * WORKFLOWS_PER_PAGE)


class DispatchTarget(NamedTuple):
    """What to dispatch for one validated ``post_merge_dispatch`` entry (#1086)."""

    #: The identifier the API accepts: the workflow file name (``ci.yml``).
    dispatch_id: str
    #: Whether the workflow declares the ``before``/``after`` range inputs (#933).
    with_range: bool


def validate_post_merge_dispatch(repo: str, names: list[str]) -> tuple[dict[str, DispatchTarget], list[str]]:
    """Each name in ``names`` must exist as a workflow in this repo and declare
    ``workflow_dispatch``. Returns ``(resolved, findings)``; non-empty
    ``findings`` blocks the merge (exit 1, no API call is fired).

    ``resolved`` maps each contract entry to the target to dispatch: the
    identifier the API accepts (the workflow FILE NAME, ``ci.yml``; #1086) and
    whether it takes the ``before``/``after`` range inputs (#933). Order matches
    ``sorted(set(names))`` so the listing is stable across rounds even when the user
    reorders the contract.
    """
    findings: list[str] = []
    if len(names) != len(set(names)):
        findings.append(
            f"post_merge_dispatch has duplicates ({sorted(set(names))}); each entry must "
            "appear once in the contract")
    seen = sorted(set(names))
    try:
        workflows = list_workflows(repo)
    except (GitHubApiError, WorkflowsListingError) as exc:
        return {}, [f"could not list the workflows to validate post_merge_dispatch: {exc}"]
    by_basename: dict[str, str] = {}
    for item in workflows:
        if not isinstance(item, dict):
            continue
        path = item.get("path")
        if isinstance(path, str):
            by_basename[_workflow_basename(path)] = path
    resolved: dict[str, DispatchTarget] = {}
    for name in seen:
        finding, with_range = _validate_one_workflow(repo, name, by_basename)
        if finding is not None:
            findings.append(finding)
            continue
        resolved[name] = DispatchTarget(_workflow_dispatch_id(by_basename[name]), with_range)
    return resolved, findings


def base_tip(repo: str, branch: str) -> str:
    """#933: punta de la rama base ANTES del merge; es el ``before`` del rango auditado.

    Se lee antes de fusionar: si no se puede leer o no es un SHA, el merge no sale (HR-34).
    """
    try:
        ref = get_json(f"repos/{repo}/git/ref/heads/{branch}")
    except GitHubApiError as exc:
        fail(f"could not read the tip of '{branch}' for the dispatch range (fail closed): {exc}", exc)
    obj = ref.get("object") if isinstance(ref, dict) else None
    sha = obj.get("sha") if isinstance(obj, dict) else None
    if not isinstance(sha, str) or not SHA_RE.match(sha):
        fail(f"the tip of '{branch}' is not a 40-char commit id: the dispatch range is not computable")
    return sha


def run_post_merge_dispatch(repo: str, items: dict[str, DispatchTarget], branch: str,
                            after: str, before: str | None) -> None:
    """Fire ``gh workflow run`` for each validated workflow on the base branch.

    Cada llamada es ``POST /repos/{repo}/actions/workflows/{workflow_id}/dispatches``
    con ``workflow_id`` = el **nombre del fichero** del workflow (``ci.yml``): el
    endpoint no reconoce el basename sin extension y responde 404, con lo que el
    detective post-merge no dispara nunca (#1086). El cuerpo es ``{"ref": branch}``:
    esa API solo acepta una rama o una etiqueta
    en ``ref`` y responde 422 a un SHA (#933), asi que el commit fusionado viaja
    solo en ``inputs.after`` para los workflows que declaran el rango. Un 2xx es
    exito (la API responde 204 sin
    cuerpo); los 4xx/5xx se acumulan y se reportan juntos al final para que
    el usuario vea el veredicto de cada entrada en una sola vuelta, no una
    a una. El merge NO se deshace en ninguno de esos fallos.
    """
    failures: list[str] = []
    for name, target in items.items():
        body: dict = {"ref": branch}
        if target.with_range:
            # #933: el detective audita `before...after`; el dispatch lo declara en `inputs`.
            body["inputs"] = {"before": before, "after": after}
        try:
            call("POST", f"repos/{repo}/actions/workflows/{target.dispatch_id}/dispatches", body)
        except GitHubApiError as exc:
            failures.append(f"{name} ({exc})")
            continue
        print(f"DISPATCHED: {name} on {branch} (after {after})")
    if failures:
        raise PostDispatchFailedError("workflows", "; ".join(failures))


def is_pull_shape(pull: object) -> bool:
    return (isinstance(pull, dict) and isinstance(pull.get("head"), dict)
            and isinstance(pull["head"].get("sha"), str) and isinstance(pull.get("labels"), list)
            and isinstance(pull.get("base"), dict) and isinstance(pull["base"].get("ref"), str)
            and isinstance(pull.get("additions"), int) and isinstance(pull.get("deletions"), int))


def compare(repo: str, left: str, right: str) -> dict:
    """Respuesta del compare de la API, validada en forma (`files` y `behind_by`)."""
    data = gh_json(f"repos/{repo}/compare/{left}...{right}")
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(files, list) or not all(
            isinstance(item, dict) and isinstance(item.get("filename"), str) for item in files):
        fail(f"compare {left}...{right} is not the API compare shape")
    return data


def paths_of(data: dict) -> set[str]:
    return {name for item in data["files"] for name in (item["filename"], item.get("previous_filename"))
            if isinstance(name, str)}


def freshness_findings(args: argparse.Namespace, pull: dict) -> list[str]:
    """La rama debe estar al dia con su base: el arbol a fusionar es el que reviso el CI.

    Con la base movida (`behind_by` > 0) el merge produce un arbol que ninguna
    ejecucion de CI vio, y el deploy lo rechaza («main moved»): se bloquea antes,
    con la accion a tomar (#787). Si ademas ambos lados cambiaron el mismo
    workflow, el merge crearia contenido de workflow nuevo y GITHUB_TOKEN no
    puede (sin permiso `workflows`, #783): GitHub respondería 403 al PUT.
    """
    base, head = pull["base"]["ref"], pull["head"]["sha"]
    ahead = compare(args.repo, base, head)
    behind = ahead.get("behind_by")
    if not isinstance(behind, int) or isinstance(behind, bool) or behind < 0:
        fail(f"compare {base}...{head} has no valid 'behind_by'")
    if behind == 0:
        return []
    mine = {name for name in paths_of(ahead) if name.startswith(WORKFLOWS_DIR)}
    shared = sorted(mine & paths_of(compare(args.repo, head, base))) if mine else []
    action = f"update the PR branch onto '{base}', let CI go green and run the governed merge again"
    if shared:
        return [f"PR #{args.pr} and '{base}' both changed {', '.join(shared)}: merging would create "
                "new workflow content, which the workflow token cannot do (no 'workflows' permission); "
                + action]
    return [f"PR #{args.pr} is behind '{base}' by {behind} commit(s): the merged tree would differ "
            "from the tree CI reviewed and the deploy evidence refuses it; " + action]


def verdict_findings(args: argparse.Namespace, sha: str, required: list[str]) -> list[str]:
    """Veredicto de los checks requeridos, esperando (acotado) a los que aun no existen.

    Un merge lanzado antes de que el CI publique `required` no es un contrato
    incumplido, es una carrera (#787): se reintenta cada `--poll-seconds` hasta
    `--wait-seconds`. Un check en fallo, o cualquier otro hallazgo, no se espera.
    """
    deadline = time.monotonic() + args.wait_seconds
    while True:
        findings = check_verdict(args.repo, args.pr, sha, required)
        waiting = [item for item in findings if item.endswith(WAITABLE)]
        if not waiting or len(waiting) != len(findings) or time.monotonic() >= deadline:
            return findings
        print(f"waiting for required checks: {'; '.join(waiting)}", file=sys.stderr)
        time.sleep(min(args.poll_seconds, max(0.0, deadline - time.monotonic())))


def gate_findings(args: argparse.Namespace, pull: dict, contract: dict,
                  required: list[str]) -> list[str]:
    body = pull.get("body") if isinstance(pull.get("body"), str) else ""
    labels = {label.get("name") for label in pull["labels"]
              if isinstance(label, dict) and isinstance(label.get("name"), str)}
    findings = []
    if pull.get("state") != "open":
        findings.append(f"PR #{args.pr} is not open (state: {pull.get('state')!r})")
    if pull.get("draft") is True:
        findings.append(f"PR #{args.pr} is a draft: the governed path merges ready PRs only")
    findings += verdict_findings(args, pull["head"]["sha"], required)
    if "chain:partial" in labels and CLOSING.search(body):
        findings.append("Closes in an intermediate PR: only the chain tip may carry closing words")
    base_ref = pull["base"]["ref"]
    if base_ref not in protected_branches(contract, None) and "chain:partial" not in labels:
        findings.append(f"PR #{args.pr} targets '{base_ref}': not a branch protected by the "
                        "contract and the PR is not a chain intermediate (chain:partial)")
    if not findings:
        findings += freshness_findings(args, pull)
    return findings


def _validate_cli_args(args: argparse.Namespace) -> None:
    """Rechaza argumentos CLI mal formados antes de cualquier llamada a red (fail closed, HR-3).

    Cada guard sale con ``fail`` (exit 2) y nunca vuelve: el caller no
    distingue entre ramas y el caller no decide si abortar; el caller
    solo llama a este helper, que aborta por sí mismo. La separación
    mantiene la API del generador de C901/PLR0912 dentro del umbral del
    ratchet shrink-only (#817).
    """
    if args.repo.count("/") != 1:
        fail("--repo must be OWNER/REPO and --pr a positive number")
    if not args.pr.isdigit() or int(args.pr) < 1:
        fail("--repo must be OWNER/REPO and --pr a positive number")
    if args.wait_seconds < 0:
        fail("--wait-seconds must be zero or more and --poll-seconds positive")
    if args.poll_seconds <= 0:
        fail("--wait-seconds must be zero or more and --poll-seconds positive")


def _resolve_post_merge_dispatch_names(contract: dict) -> list[str]:
    """Resuelve ``contract["post_merge_dispatch"]`` en una lista de basenames.

    Ausente o lista vacía: warning a stderr y return ``[]``. Lista mal
    tipada: fail closed. Lista no vacía: valida que cada item sea una
    cadena no vacía y devuelve los nombres. La lista validada vive en
    ``.github/host-contract.json::post_merge_dispatch`` del consumer
    y cada nombre debe existir en el repo y declarar ``on:
    workflow_dispatch:`` (esa verificación vive en
    ``validate_post_merge_dispatch``).
    """
    raw = contract.get("post_merge_dispatch")
    if isinstance(raw, list):
        if not raw:
            print("WARNING: post_merge_dispatch is an empty list; post-merge dispatch "
                  "skipped (#817).", file=sys.stderr)
            return []
        if not all(isinstance(item, str) and item for item in raw):
            fail("post_merge_dispatch must be a list of non-empty strings")
        return [item for item in raw if isinstance(item, str)]
    if "post_merge_dispatch" in contract:
        fail("post_merge_dispatch must be a list of non-empty strings when declared")
    print("WARNING: post_merge_dispatch is not declared; post-merge dispatch "
          "skipped (#817, "
          "https://github.com/DysTelefonica/team-skills/issues/817).",
          file=sys.stderr)
    return []





def read_commit_tree(repo: str, sha: str) -> str:
    """Arbol de un commit por la API; falla cerrado si la forma no vale (#1265).

    Lee ``commit.tree.sha`` (la forma real de la API, #928) con el lector
    compartido de ``check_push_compliance``: una sola implementacion de «el
    arbol de un commit». Si la API falla, ``gh_json`` ya falla cerrado; si la
    forma no trae el arbol, tambien: un commit ilegible nunca se reporta como
    coincidencia.
    """
    commit = gh_json(f"repos/{repo}/commits/{sha}")
    tree = commit_tree_sha(commit)
    if tree is None:
        fail(f"the commit {sha} does not expose 'commit.tree.sha': the integrated tree "
             "cannot be measured (fail closed)")
    return tree


def _verdict_summary(merge_sha: str, merge_tree: str, head_sha: str, head_tree: str,
                     drift: bool) -> str:
    """El veredicto en markdown para el resumen del job (#1265)."""
    verdict = VERDICT_DRIFT if drift else VERDICT_MATCH
    return ("## Integrated tree evidence\n"
            "\n"
            f"- merge commit: `{merge_sha}`\n"
            f"- merge tree: `{merge_tree}`\n"
            f"- verified head: `{head_sha}`\n"
            f"- head tree: `{head_tree}`\n"
            f"- verdict: {verdict}\n")


def _not_measured_summary(cause: str) -> str:
    """El resumen del job cuando la medicion no se pudo hacer (#1279).

    Publica la causa en vez de callarse: el resumen no afirma que el arbol
    coincide —no se ha medido— pero dice por que no se pudo medir.
    """
    return ("## Integrated tree evidence\n"
            "\n"
            f"- verdict: {NOT_MEASURED}: {cause}\n")


def _write_step_summary(text: str) -> None:
    """Publica el veredicto en ``$GITHUB_STEP_SUMMARY`` (#1265).

    Sin la variable (corridas locales, tests) el veredicto queda solo en
    stdout: no es un fallo. Con la variable definida, si la escritura falla el
    job falla cerrado (2): no se puede afirmar que se publico la evidencia.
    """
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not path:
        return
    try:
        with Path(path).open("a", encoding="utf-8") as handle:
            handle.write(text)
    except OSError as exc:
        fail(f"the integrated-tree verdict could not be written to GITHUB_STEP_SUMMARY: {exc}",
             exc)


def _publish_not_measured(cause: str) -> None:
    """Publica la causa de una medicion fallida (#1279).

    Si ni esta escritura puede, no hay mas que hacer: la causa queda en
    stdout y el job sale en rojo al final. Esto nunca interrumpe el trabajo
    post-merge.
    """
    try:
        _write_step_summary(_not_measured_summary(cause))
    except FailClosedError as exc:
        print(f"INTEGRATED TREE: the cause could not be published to the job summary: {exc}",
              file=sys.stderr)


def integrated_tree_verdict(repo: str, merge_sha: str, head_sha: str) -> tuple[bool, str | None]:
    """Mide el arbol integrado y publica lo que pueda; devuelve (drift, causa).

    La causa no es None cuando la medicion no se pudo hacer o su veredicto no
    se pudo publicar. Esta funcion NO falla cerrado: el merge ya esta hecho
    (#1265) y un fallo de medicion no puede impedir cerrar las issues ni
    despachar los workflows post-merge (#1279). El job sale en rojo al final,
    con 2, nombrando la causa.

    Se mide INMEDIATAMENTE despues del merge: la evidencia se publica aunque
    un paso posterior (cerrar la issue, despachar) falle. Compara ARBOLES y no
    commits: con la base contenida en el head, merge, squash y rebase dan el
    mismo arbol, asi que el control no depende del metodo de merge por defecto
    del host.

    El veredicto nombra los dos commits y los dos arboles: lo que se lee en el
    log es la misma evidencia que queda en el resumen del job. No se crea ni
    muta ningun check-run: nada en la cabeza del PR parece una segunda
    evaluacion del mismo trabajo.
    """
    try:
        merge_tree = read_commit_tree(repo, merge_sha)
        head_tree = read_commit_tree(repo, head_sha)
    except FailClosedError as exc:
        cause = str(exc)
        print(f"INTEGRATED TREE: {NOT_MEASURED}: {cause}")
        _publish_not_measured(cause)
        return False, cause
    drift = merge_tree != head_tree
    verdict = VERDICT_DRIFT if drift else VERDICT_MATCH
    print(f"INTEGRATED TREE: merge commit {merge_sha} tree {merge_tree} | "
          f"verified head {head_sha} tree {head_tree} | verdict: {verdict}")
    try:
        _write_step_summary(_verdict_summary(merge_sha, merge_tree, head_sha, head_tree, drift))
    except FailClosedError as exc:
        return drift, str(exc)
    return drift, None


def _verdict_code(drift: bool, measure_error: str | None) -> int:
    """El codigo final tras el cierre y el dispatch: 2, 5 o 0 (#1279).

    La precedencia es 3 (cierre fallido) y 4 (dispatch fallido), que se
    devuelven antes de llegar aqui; despues 2 por medicion no hecha y
    despues 5 por drift. Un fallo de medicion no puede esconder un drift
    que ya se midio: el veredicto queda en la linea de stdout.
    """
    if measure_error is not None:
        print(f"MERGED but the integrated tree was not measured: {measure_error}", file=sys.stderr)
        return 2
    if drift:
        print("MERGED but not the tested tree: the integrated tree is NOT the tree CI reviewed",
              file=sys.stderr)
        return INTEGRATED_TREE_DRIFT
    return 0


def _merge_close_and_dispatch(args: argparse.Namespace, pull: dict, contract: dict,
                              required: list[str], dispatch_names: list[str]) -> int:
    """Gate -> merge -> evidencia del arbol -> close -> post-merge dispatch.

    Devuelve el código de salida: 1 (gate bloquea), 2 (fail closed o
    medición no hecha, #1279), 3
    (close failed), 4 (dispatch failed), 5 (el arbol integrado no es el del
    head verificado, #1265) o 0 (success). El merge ya
    hecho no se deshace en ninguno de los caminos 2, 3, 4 y 5 (idempotencia
    hacia atrás, HR-34). Un fallo de la medición del arbol no salta el cierre
    ni el dispatch: la causa se guarda y se devuelve al final (#1279).
    """
    findings = gate_findings(args, pull, contract, required)
    resolved: dict[str, DispatchTarget] = {}
    if dispatch_names:
        resolved, dispatch_findings = validate_post_merge_dispatch(args.repo, dispatch_names)
        findings.extend(dispatch_findings)
    if findings:
        for finding in findings:
            print(f"  BLOCKED  {finding}")
        print(f"VERDICT: PR #{args.pr} not merged: contract not satisfied")
        return 1
    # #933: el `before` del rango es la punta de la base antes del merge. Solo
    # se lee si algun workflow lo recibe: `DispatchTarget` es una tupla y
    # `any(resolved.values())` seria siempre cierto (#1086).
    before = (base_tip(args.repo, pull["base"]["ref"])
              if any(target.with_range for target in resolved.values()) else None)
    sha = gh_merge(args.repo, args.pr, pull["head"]["sha"])
    print(f"MERGED: PR #{args.pr} merged by the governed path as {sha}")
    # #1265: el veredicto del arbol integrado se mide y se publica INMEDIATAMENTE
    # despues del merge, antes de cerrar y antes del dispatch: la evidencia sale
    # aunque un paso posterior falle.
    # #1279: un fallo de la medicion no interrumpe ese trabajo: la causa se
    # guarda y se devuelve al final, con 2. El merge ya esta hecho.
    drift, measure_error = integrated_tree_verdict(args.repo, sha, pull["head"]["sha"])
    try:
        close_linked_issues(args.repo, pull, args.pr, sha)
    except (MergedNotClosedError, FailClosedError) as exc:
        print(f"MERGED but not closed: {exc}", file=sys.stderr)
        return 3
    # #817: post-merge dispatch via workflow_dispatch (que ``GITHUB_TOKEN``
    # sí dispara). Cada workflow validado se lanza sobre la rama base; el SHA
    # fusionado viaja en ``inputs.after`` para los que declaran el rango (#933).
    if resolved:
        try:
            run_post_merge_dispatch(args.repo, resolved, pull["base"]["ref"], sha, before)
        except PostDispatchFailedError as exc:
            print(f"MERGED but not dispatched: {exc}", file=sys.stderr)
            return POST_DISPATCH_FAILED
    return _verdict_code(drift, measure_error)


def run(args: argparse.Namespace) -> int:
    """Punta de la cadena (#817): orquesta validación + merge + close + dispatch.

    La orquestación se delega a ``_merge_close_and_dispatch``; ``run()`` se
    queda en 2 ramas (load contract + is_pull_shape) para mantenerse por
    debajo del umbral del ratchet shrink-only de ruff extendido (C901/PLR0912).
    """
    _validate_cli_args(args)
    contract = load_json(args.contract, "contract")
    if not isinstance(contract, dict):
        fail("contract must be a JSON object")
    required = required_checks(contract)
    dispatch_names = _resolve_post_merge_dispatch_names(contract)
    pull = gh_json(f"repos/{args.repo}/pulls/{args.pr}")
    if not is_pull_shape(pull):
        fail(f"PR #{args.pr} is not the API pull shape")
    return _merge_close_and_dispatch(args, pull, contract, required, dispatch_names)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="governed merge vs the host contract")
    parser.add_argument("--repo", required=True, help="OWNER/REPO")
    parser.add_argument("--pr", required=True, help="pull request number")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--max-lines", type=int, default=None,
                        help="deprecated and ignored: pr-size owns the review budget (#787)")
    parser.add_argument("--wait-seconds", type=float, default=DEFAULT_WAIT,
                        help="how long to wait for required checks not yet published")
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL)
    args = parser.parse_args()
    try:
        code = run(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
