#!/usr/bin/env python3
# ci-pattern asset — post-push compliance detective (DysTelefonica/team-skills#303)
"""Detective post-push: audita cada push a una rama protegida por contrato.

El contrato declara `protected_branches` y `required_checks`; para cada
commit del evento `push` la API debe mostrar un PR mergeado y los checks
requeridos en verde sobre la cabeza de ese PR. Con la clave opt-in
`governed_path` el detective exige además que el merge haya llegado por
el camino gobernado (#909): el fusionador en la lista de actores o la
firma del commit de merge. Una violación abre una issue
de incidente con la evidencia (SHA, autor, checks) y sale con 1; un push
conforme sale con 0. Entorno insalvable falla cerrado con 2 (HR-3): nunca
verde sin medición. Runner-enforced (HR-34): sin protección de rama este
detective es el control que actúa y deja evidencia.

Ventana de settle (#1019): un push llega segundos después del merge y el check
requerido suele ser el rollup final de un DAG, así que la API todavía no lo ha
publicado. `missing` y `pending` son el CI en marcha, no un incumplimiento: se
espera `--poll-seconds` hasta `--wait-seconds` (por defecto 900 s). El workflow
del consumer debe declarar `timeout-minutes` mayor que `--wait-seconds`; si el
techo se agota sin veredicto, el estado es indeterminado y falla cerrado sin
abrir incidente.
"""
# #1019: las anotaciones se difieren. `settled_verdict` anota `push: PushAudit`
# y se define antes que la clase: con evaluacion eager (Python < 3.14, el que
# corre el CI) eso es un NameError al importar, y como check_governed_merge
# importa este modulo, tumbaba tambien la suite del gate pre-merge.
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn
from urllib.parse import urlencode

from github_rest import GitHubApiError, call, get_json

SHA_RE = re.compile(r"^[0-9a-f]{40}$")

# #1019: un check que la API aun no ha publicado, o que aun no ha terminado, no
# es un incumplimiento: es el CI en marcha. Son los mismos dos veredictos que el
# gate pre-merge marca como esperables (`check_governed_merge.WAITABLE`).
WAITABLE = (": missing", ": pending")
# El check requerido suele ser el rollup final de un DAG: en APAP_WEB tardo hasta
# 13 min desde el merge (#1019). El techo cubre esa observacion; el workflow del
# consumer debe declarar `timeout-minutes` mayor que el (ver SKILL.md).
DEFAULT_WAIT, DEFAULT_POLL = 900.0, 30.0


class FailClosedError(Exception):
    """A condition the detective refuses to interpret; maps to exit 2."""


def fail(message: str, cause: BaseException | None = None) -> NoReturn:
    """Eleva la condición de fallo cerrado; el mensaje lo fija quien la detecta."""
    raise FailClosedError(message) from cause


def load_json(path: str, what: str) -> object:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except OSError as exc:
        fail(f"unreadable {what} '{path}': {exc}", exc)
    except json.JSONDecodeError as exc:
        fail(f"{what} '{path}' is not valid JSON: {exc}", exc)


def gh_json(path: str) -> object:
    try:
        return get_json(path)
    except GitHubApiError as exc:
        fail(f"GitHub API {path} failed (fail closed): {exc}", exc)


def gh_open_incident(repo: str, title: str, body: str) -> str:
    try:
        text, _ = call("POST", f"repos/{repo}/issues", {"title": title, "body": body})
    except GitHubApiError as exc:
        fail(f"the incident issue could not be opened (fail closed): {exc}", exc)
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        fail(f"the incident issue response is not valid JSON: {exc}", exc)
    url = data.get("html_url") if isinstance(data, dict) else None
    if not isinstance(url, str) or not url:
        fail("the incident issue response carries no html_url")
    return url


def audit_identity(event: dict) -> tuple[str, str]:
    """OWNER/REPO y rama del evento; un evento sin forma es duda (HR-3)."""
    if not isinstance(event.get("ref"), str) or not event["ref"].startswith("refs/heads/"):
        fail("event has no 'refs/heads/' ref to audit")
    repo = event.get("repository")
    if not isinstance(repo, dict) or not isinstance(repo.get("full_name"), str) \
            or repo["full_name"].count("/") != 1:
        fail("event has no OWNER/REPO 'repository.full_name'")
    return repo["full_name"], event["ref"][len("refs/heads/"):]


def push_range(event: dict) -> tuple[str, str]:
    """#933: rango `before...after` de un push o de un `workflow_dispatch` (en `inputs`).

    Un dispatch no trae `before`/`after` de nivel superior: el lanzador los declara en
    `inputs`. Sin ellos el rango no es auditable y el detective falla cerrado (HR-3).
    """
    source = event.get("inputs", event)
    if not isinstance(source, dict):
        fail("workflow_dispatch 'inputs' must be an object carrying 'before' and 'after'")
    before, after = source.get("before"), source.get("after")
    if not isinstance(before, str) or not SHA_RE.match(before) \
            or not isinstance(after, str) or not SHA_RE.match(after):
        fail("event 'before'/'after' are not 40-char commit ids")
    return before, after


def audit_event(event: dict) -> tuple[str, list[dict]]:
    repo_name, _ = audit_identity(event)
    before, after = push_range(event)
    if before == "0" * 40:
        fail("before of zeros (new branch): the pushed range is not auditable")
    if event.get("deleted") is True:
        return repo_name, []  # el borrado lo trata el control de force-push
    if "inputs" in event:
        # #933: un dispatch no trae la lista de commits; audit() la completa con compare_commits.
        return repo_name, [{"id": after}]
    commits = event.get("commits")
    if not isinstance(commits, list):
        fail("event has no 'commits' list to audit")
    head = event.get("head_commit")
    if isinstance(head, dict) and isinstance(head.get("id"), str) and all(
            c.get("id") != head["id"] for c in commits if isinstance(c, dict)):
        commits = commits + [head]
    if not commits:
        fail("empty pushed range: no commits to audit")
    audited: list[dict] = []
    for commit in commits:
        if not isinstance(commit, dict) or not isinstance(commit.get("id"), str) \
                or not SHA_RE.match(commit["id"]):
            fail("a pushed commit carries no 40-char id")
        if all(c["id"] != commit["id"] for c in audited):
            audited.append(commit)
    return repo_name, audited


def protected_branches(contract: dict, default_branch: object) -> list[str]:
    declared = contract.get("protected_branches")
    if declared is None:
        return [default_branch] if isinstance(default_branch, str) and default_branch else []
    if not isinstance(declared, list) or not declared or any(
            not isinstance(name, str) or not name for name in declared):
        fail("contract 'protected_branches' must be a non-empty list of branch names")
    return declared


def required_checks(contract: dict) -> list[str]:
    entries = contract.get("required_checks", [])
    if not isinstance(entries, list):
        fail("contract 'required_checks' must be a list")
    names: list[str] = []
    for entry in entries:
        if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) \
                or not entry["name"]:
            fail("every contract required check needs a non-empty 'name'")
        if entry["name"] not in names:
            names.append(entry["name"])
    return names


def tracker_branch(contract: dict) -> str | None:
    """#333 (regla 7): con `tracker_branch` declarado, lo que llega a una rama
    protegida tiene que venir del PR del tracker (modelo «rama de feature con PR
    tracker en borrador»). Sin declaración el control no cambia: una cadena
    apilada hacia main no necesita tracker, y exigirlo sin declararlo
    inventaria una regla que el contrato no pide."""
    value = contract.get("tracker_branch")
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        fail("contract 'tracker_branch' must be a non-empty string")
    return value.strip()


def governed_path_config(contract: dict) -> dict | None:
    """Clave opt-in `governed_path` del contrato del host (#909).

    Sin declaración (o con `enabled: false`) el detective no cambia: un
    repo que no declara el camino gobernado no paga un control que no
    pidió. Con `enabled: true`, un PR mergeado en una rama protegida
    cumple si su FUSIONADOR está en `actors` o si el commit de merge
    habla la firma gobernada; si no, es una violación. Un contrato
    inservible falla cerrado (HR-3): duda, nunca verde.

    Las dos rutas no valen lo mismo: la del actor es la FUERTE (la
    plataforma firma `merged_by` y nadie lo edita desde la interfaz) y la
    del mensaje es la DEBIL (quien mergea por la web puede escribir el
    título del commit de merge y hacer pasar su merge por gobernado). Este
    control detecta un bypass ACCIDENTAL, no uno deliberado: quien quiera
    engañarlo tiene la ruta débil a mano. Para que el camino gobernado sea
    el único con valor, declara en `actors` los fusionadores legítimos y
    lee `merge_marker` como corroboración, no como prueba."""
    cfg = contract.get("governed_path")
    if cfg is None:
        return None
    if not isinstance(cfg, dict):
        fail("contract 'governed_path' must be an object")
    enabled = cfg.get("enabled", False)
    if not isinstance(enabled, bool):
        fail("contract 'governed_path.enabled' must be a boolean")
    if not enabled:
        return None
    actors = cfg.get("actors")
    if not isinstance(actors, list) or not actors or any(
            not isinstance(actor, str) or not actor.strip() for actor in actors):
        fail("contract 'governed_path.actors' must be a non-empty list of non-empty strings")
    marker_source = cfg.get("merge_marker", r"^Merge PR #\d+ \(governed merge\)$")
    if not isinstance(marker_source, str):
        fail("contract 'governed_path.merge_marker' must be a string")
    try:
        marker = re.compile(marker_source)
    except re.error as exc:
        fail(f"contract 'governed_path.merge_marker' is not a valid regex: {exc}", exc)
    return {"actors": list(actors), "marker": marker}


def check_governed_path(cfg: dict, pull: dict, repo: str) -> list[str]:
    """Verificar que el merge llegó por el camino gobernado.

    El fusionador se lee del DETALLE del PR (`GET /repos/{repo}/pulls/{n}`
    -> `merged_by`): la lista de PRs de un commit trae `user.login`, que
    es el AUTOR del PR, no quien lo fusionó. La ruta del actor se resuelve
    antes que la del mensaje: cuando el fusionador está en la lista, no
    se hace la llamada extra al commit de merge."""
    number = pull.get("number")
    merged_by = pull.get("merged_by")
    login = merged_by.get("login") if isinstance(merged_by, dict) else None
    actors = cfg["actors"]
    if isinstance(login, str) and login in actors:
        return []  # ruta 1: el fusionador está en la lista de actores
    sha = pull.get("merge_commit_sha")
    if not isinstance(sha, str) or not SHA_RE.match(sha):
        return [f"PR #{number} merged by '{login or 'unknown'}' is outside the governed path "
                f"(actors: {', '.join(actors)}) and carries no merge commit to read the marker"]
    commit = gh_json(f"repos/{repo}/commits/{sha}")
    message = ""
    if isinstance(commit, dict) and isinstance(commit.get("commit"), dict):
        raw = commit["commit"].get("message")
        message = raw if isinstance(raw, str) else ""
    first_line = message.split("\n", 1)[0]
    if cfg["marker"].match(first_line):
        return []  # ruta 2: el commit de merge habla el idioma gobernado
    return [f"PR #{number} merged by '{login or 'unknown'}' is outside the governed path "
            f"(actors: {', '.join(actors)}; merge commit first line: "
            f"'{first_line or '<empty>'}')"]


def check_tree_integrity(repo: str, detail: dict, commit_sha: str) -> list[str]:
    """Invariante de integridad de arbol (#426 safe-merge, #928).

    Se comprueba solo si `commit_sha` es el `merge_commit_sha` del PR y la base
    en el momento del merge (primer padre de ese commit) es ancestro del head,
    es decir, `compare base...head` sale `ahead` o `identical`. En ese caso el
    arbol del commit de merge debe ser el arbol del head revieweado.

    No se comprueba en estos casos (no son violaciones):
    - commits intermedios de la PR: su arbol nunca es el del head;
    - main avanzo desde la ultima actualizacion de la rama: el compare sale
      `diverged` o `behind` y el arbol del merge difiere legitimamente del head;
    - rebase-merge (LIMITACION CONOCIDA): `merge_commit_sha` es el ultimo commit
      rebasado y su primer padre es el rebasado anterior, no la base. El compare
      sale `diverged` y el chequeo no aplica, aunque haya drift. Queda fijado por
      `test_limitation_rebase_merge_drift_is_not_detected`.

    Datos ausentes o de forma inesperada no bloquean (fall-open), como antes.
    """
    head_sha = detail["head"]["sha"]
    if commit_sha != detail.get("merge_commit_sha"):
        return []
    merge = gh_json(f"repos/{repo}/commits/{commit_sha}")
    base = _first_parent(merge)
    if base is None:
        return []
    comparison = gh_json(f"repos/{repo}/compare/{base}...{head_sha}")
    if not isinstance(comparison, dict) or comparison.get("status") not in ("ahead", "identical"):
        return []
    merge_tree = commit_tree_sha(merge)
    head_tree = commit_tree_sha(gh_json(f"repos/{repo}/commits/{head_sha}"))
    if merge_tree is None or head_tree is None or merge_tree == head_tree:
        return []
    return [
        f"arbol de PR #{head_tree[:12]} vs merge {commit_sha[:7]}: difieren en el arbol; "
        "esto indica que el arbol mergeado no es identico al arbol revieweado. "
        "Verifique que se uso safe-merge (#426)."
    ]


def _first_parent(commit: object) -> str | None:
    """Primer padre de un commit de la API (la base en un merge commit)."""
    parents = commit.get("parents") if isinstance(commit, dict) else None
    if not isinstance(parents, list) or not parents or not isinstance(parents[0], dict):
        return None
    sha = parents[0].get("sha")
    return sha if isinstance(sha, str) and SHA_RE.match(sha) else None


def commit_tree_sha(commit: object) -> str | None:
    """Arbol de un commit de la API: la forma real es commit.tree.sha (#928).

    Lector COMPARTIDO (#1265): lo usan el detective post-push
    (``check_tree_integrity``) y el merge gobernado
    (``check_governed_merge``), para que no haya dos implementaciones de «el
    arbol de un commit». ``None`` cuando la forma no vale: el que llama decide
    si eso es duda (el detective falla abierto) o fallo cerrado (el veredicto
    del merge, que nunca informa de una coincidencia que no ha medido).
    """
    tree = ((commit.get("commit") or {}).get("tree") or {}).get("sha") if isinstance(commit, dict) else None
    return tree if isinstance(tree, str) and SHA_RE.match(tree) else None


def check_verdict(repo: str, pr_number: object, head_sha: str, required: list[str]) -> list[str]:
    # Politica de verificacion (#326): se leen las check-runs con
    # filter=latest sobre la primera pagina (per_page=100); la API legada de
    # commit statuses se ignora a proposito. Ambos limites sesgan a FALSO
    # POSITIVO (un check requerido ausente de la pagina se reporta missing y
    # bloquea), nunca a falso verde: un gate que no mide no puede aprobar.
    runs = gh_json(f"repos/{repo}/commits/{head_sha}/check-runs?filter=latest&per_page=100")
    if not isinstance(runs, dict) or not isinstance(runs.get("check_runs"), list):
        fail(f"check-runs for PR #{pr_number} is not the API object shape")
    latest: dict = {}
    for run in runs["check_runs"]:
        if not isinstance(run, dict) or not isinstance(run.get("name"), str):
            continue
        if run["name"] not in latest or run.get("id", 0) > latest[run["name"]].get("id", 0):
            latest[run["name"]] = run
    verdicts = []
    for name in required:
        run = latest.get(name)
        if run is None:
            verdicts.append(f'required check "{name}": missing')
            continue
        conclusion = run.get("conclusion")
        if run.get("status") == "completed" and conclusion == "success":
            continue
        verdict = f'required check "{name}": {conclusion or "pending"}'
        # #1019: una conclusion terminal es un hecho consumado y es la evidencia:
        # citarla para que la incidencia nombre el run y el job, no solo la palabra.
        url = run.get("html_url")
        if run.get("status") == "completed" and isinstance(url, str) and url:
            verdict = f"{verdict} [{url}]"
        verdicts.append(verdict)
    return verdicts


def _is_waitable(finding: str) -> bool:
    """El hallazgo dice que el CI aun no publico o no termino el check (#1019)."""
    return finding.endswith(WAITABLE)


def settled_verdict(push: PushAudit, repo: str, pr_number: object,
                    head_sha: str, required: list[str]) -> list[str]:
    """Veredicto de los checks requeridos, esperando (acotado) a los aun no publicados.

    Un push a una rama protegida llega segundos despues del merge, y el check
    requerido es el rollup final de un DAG: a esa altura la API todavia no lo ha
    publicado. Juzgarlo `missing` acusa un incumplimiento que solo esta en marcha,
    y el falso positivo es deterministico, no una carrera (#1019; medido en
    ardelperal/APAP_WEB#2280, donde el rollup quedo verde 13 min despues). Se
    reintenta cada `--poll-seconds` hasta `--wait-seconds` mientras TODOS los
    hallazgos sean esperables; una conclusion terminal no se espera.

    Agotado el techo con el check sin publicar, el estado es duda (HR-3), no una
    infraccion: falla cerrado y NO abre incidente, porque acusar seria repetir el
    defecto que este gate corrige.
    """
    deadline = time.monotonic() + push.wait_seconds
    while True:
        findings = check_verdict(repo, pr_number, head_sha, required)
        pending = [item for item in findings if _is_waitable(item)]
        if not pending or len(pending) != len(findings) or time.monotonic() >= deadline:
            break
        print(f"waiting for required checks: {'; '.join(pending)}", file=sys.stderr)
        time.sleep(min(push.poll_seconds, max(0.0, deadline - time.monotonic())))
    if findings and all(_is_waitable(item) for item in findings):
        fail(f"PR #{pr_number} (head {head_sha}): tras {push.wait_seconds:.0f}s el CI no publico "
             f"ni termino {'; '.join(findings)} — estado INDETERMINADO, no incumplimiento: el "
             f"techo no cubrio el rollup. Sube --wait-seconds y el `timeout-minutes` del "
             f"workflow por encima del (ver SKILL.md)")
    return findings


def compare_range(repo: str, before: str, after: str) -> list[dict]:
    """Rango completo según el compare (la API devuelve hasta 250; #326)."""
    compare = gh_json(f"repos/{repo}/compare/{before}...{after}")
    if not isinstance(compare, dict) or not isinstance(compare.get("total_commits"), int):
        fail(f"compare for {before}...{after} is not the API compare shape")
    range_commits = compare.get("commits")
    if not isinstance(range_commits, list):
        fail("the compare response has no 'commits' list to audit")
    # El payload de un push trae como maximo 20 commits (#326): el
    # rango completo se audita con la lista 'commits' del compare
    # (la API devuelve hasta 250). Solo un total que SUPERA lo que
    # el compare devuelve deja el rango sin auditar entero: duda,
    # exit 2 (HR-3) — un payload corto ya auditable no es un rojo.
    if compare["total_commits"] > len(range_commits):
        fail(
            f"the compare reports {compare['total_commits']} commits but returned "
            f"{len(range_commits)}: the pushed range exceeds what the API returns "
            "and cannot be audited in full")
    commits = []
    for entry in range_commits:
        sha = entry.get("sha") if isinstance(entry, dict) else None
        raw = entry.get("commit") if isinstance(entry, dict) else None
        commit = raw if isinstance(raw, dict) else {}
        author = commit.get("author") if isinstance(commit.get("author"), dict) else {}
        if not isinstance(sha, str) or not SHA_RE.match(sha):
            fail("the compare returned a commit without a 40-char sha")
        commits.append({"id": sha, "author": {"name": author.get("name")}})
    return commits


def compare_commits(event: dict, repo: str) -> list[dict]:
    """Rango completo del push según el compare (la API devuelve hasta 250; #326)."""
    before, after = push_range(event)
    return compare_range(repo, before, after)


def associated_pulls(repo: str, sha: str) -> list[dict]:
    """PRs asociadas a un commit; la API devuelve una lista JSON."""
    prs = gh_json(f"repos/{repo}/commits/{sha}/pulls")
    if not isinstance(prs, list):
        fail(f"pulls for commit {sha} is not a JSON array")
    return [pull for pull in prs if isinstance(pull, dict)]


def pull_detail(repo: str, number: object) -> dict:
    """Detalle de una PR con la forma que exige el detective."""
    detail = gh_json(f"repos/{repo}/pulls/{number}")
    if not isinstance(detail, dict) or not isinstance(detail.get("head"), dict) \
            or not isinstance(detail["head"].get("sha"), str):
        fail(f"PR #{number} detail is not the API pull shape")
    return detail


def listed_pulls(repo: str, branch: str) -> list[dict]:
    """PRs cerradas con la rama como base; la API devuelve una lista JSON."""
    query = urlencode({"state": "closed", "base": branch, "per_page": 100,
                       "sort": "updated", "direction": "desc"})
    pulls = gh_json(f"repos/{repo}/pulls?{query}")
    if not isinstance(pulls, list):
        fail(f"pulls listing for {repo} is not a JSON array")
    return [pull for pull in pulls if isinstance(pull, dict)]


def pull_covers_push(pull: dict, branch: str, after: str) -> bool:
    """La PR esta fusionada en la rama y su merge commit o head es la punta."""
    if not (pull.get("merged") or pull.get("merged_at")):
        return False
    base = pull.get("base")
    head = pull.get("head")
    head_sha = head.get("sha") if isinstance(head, dict) else None
    if not isinstance(base, dict) or base.get("ref") != branch:
        return False
    return pull.get("merge_commit_sha") == after or head_sha == after


def merged_push_pulls(repo: str, branch: str, after: str) -> list[dict]:
    """PRs fusionadas en la rama cuyo merge commit o head es la punta del push."""
    candidates = [pull for pull in associated_pulls(repo, after)
                  if pull_covers_push(pull, branch, after)]
    if candidates:
        return candidates
    return [pull for pull in listed_pulls(repo, branch)
            if pull_covers_push(pull, branch, after)]


def merge_coverage(repo: str, branch: str, after: str) -> tuple[set[str], list[dict]]:
    """Commits de base..merge cubiertos por la PR que fusionó la punta del push."""
    covered = set()
    covering = []
    for pull in merged_push_pulls(repo, branch, after):
        number = pull.get("number")
        detail = pull_detail(repo, number)
        base = detail.get("base")
        head = detail.get("head")
        base_sha = base.get("sha") if isinstance(base, dict) else None
        head_sha = head.get("sha") if isinstance(head, dict) else None
        merge_commit_sha = detail.get("merge_commit_sha")
        merge_sha = merge_commit_sha
        if not isinstance(merge_sha, str) or not SHA_RE.match(merge_sha):
            merge_sha = head_sha
        if not isinstance(merge_sha, str) or not SHA_RE.match(merge_sha):
            fail(f"PR #{number} does not expose a valid base..merge range")
        if not isinstance(base_sha, str) or not SHA_RE.match(base_sha):
            # El primer padre del commit de merge es la base efectiva; sirve
            # cuando el detalle de la PR no trae `base.sha`.
            if isinstance(merge_commit_sha, str) and SHA_RE.match(merge_commit_sha):
                base_sha = _first_parent(gh_json(f"repos/{repo}/commits/{merge_commit_sha}"))
            if not isinstance(base_sha, str) or not SHA_RE.match(base_sha):
                fail(f"PR #{number} does not expose a valid base..merge range")
        if detail.get("merge_commit_sha") != after and head_sha != after:
            fail(f"PR #{number} does not match the pushed commit")
        covered.update(commit["id"] for commit in compare_range(repo, base_sha, merge_sha))
        covering.append(pull)
    return covered, covering


@dataclass(frozen=True)
class PushAudit:
    """Contexto fijo de la auditoría de un push: lo que no cambia de un commit a otro."""

    repo: str
    branch: str
    event: dict
    after: str
    tracker: str | None
    required: list[str]
    governed: dict | None
    seen_prs: set
    covered: set[str]
    covering_pulls: tuple[dict, ...]
    wait_seconds: float
    poll_seconds: float


def commit_findings(push: PushAudit, commit: dict) -> list[str]:
    repo, branch, event, tracker, required, governed, seen_prs = (
        push.repo, push.branch, push.event, push.tracker, push.required,
        push.governed, push.seen_prs)
    author = commit.get("author") if isinstance(commit.get("author"), dict) else {}
    author_name = author.get("name") or event.get("pusher", {}).get("name") or "unknown author"
    if commit["id"] in push.covered and commit["id"] != push.after:
        return []
    prs = associated_pulls(repo, commit["id"])
    known = {pull.get("number") for pull in prs}
    prs += [pull for pull in push.covering_pulls if pull.get("number") not in known]
    merged = [p for p in prs if p.get("merged") or p.get("merged_at")]
    into_branch = [p for p in merged
                   if isinstance(p.get("base"), dict) and p["base"].get("ref") == branch]
    if not into_branch:
        reason = ("no merged PR into "
                  f"{branch} (merged elsewhere does not cover this push)"
                  if merged else "no merged PR associated (push outside the governed merge path)")
        return [f"commit {commit['id']} — author: {author_name}\n  reason: {reason}"]
    if tracker is not None:
        heads = sorted({p["head"]["ref"] for p in into_branch
                        if isinstance(p.get("head"), dict)
                        and isinstance(p["head"].get("ref"), str)})
        if tracker not in heads:
            return [
                f"commit {commit['id']} — author: {author_name}\n"
                f"  reason: el contrato declara tracker_branch '{tracker}' y ningún PR "
                f"mergeado en {branch} tiene esa rama como cabeza "
                f"(cabezas: {heads or 'desconocidas'}) — lo que llega a una rama "
                "protegida tiene que venir del PR del tracker (#333, regla 7)"]
    findings: list[str] = []
    for pull in into_branch:
        number = pull.get("number")
        # #928: el commit de merge puede llegar despues de que un commit intermedio
        # haya consumido el PR; entonces solo se audita el arbol del merge.
        is_merge = pull.get("merge_commit_sha") == commit["id"]
        if number in seen_prs and not is_merge:
            continue
        first = number not in seen_prs
        seen_prs.add(number)
        detail = pull_detail(repo, number)
        # El fusionador sale del DETALLE del PR: `pull["user"]["login"]` (la
        # lista de PRs de un commit) es el AUTOR del PR, no quien lo fusionó.
        merged_by = detail.get("merged_by")
        login = merged_by.get("login") if isinstance(merged_by, dict) else None
        author_login = (pull.get("user") or {}).get("login") or "unknown"
        verdicts = settled_verdict(push, repo, number, detail["head"]["sha"], required) if first else []
        # Verificacion del arbol: solo para el commit de merge (ver check_tree_integrity).
        verdicts += check_tree_integrity(repo, detail, commit["id"])
        if governed is not None and first:
            verdicts += check_governed_path(governed, detail, repo)
        if verdicts:
            evidence = "\n".join(f"  {line}" for line in verdicts)
            findings.append(f"commit {commit['id']} — author: {author_name} — PR #{number} "
                            f"merged by {login or 'unknown'} (PR author: {author_login}; "
                            f"head {detail['head']['sha']})\n{evidence}")
    return findings


def incident_report(event: dict, branch: str, commits: list[dict],
                    findings: list[str]) -> tuple[str, str]:
    before, after = push_range(event)  # #933: un dispatch no trae before/after de primer nivel
    title = f"incident: push compliance violation on '{branch}' ({commits[0]['id'][:7]})"
    body = ("Post-push detective (`check_push_compliance`), rama protegida por contrato: "
            f"{branch}\n\npusher: {event.get('pusher', {}).get('name', 'unknown')}\n"
            f"rango: {before}..{after}\n\n"
            + "\n\n".join(findings)
            + "\n\nTodo cambio entra en main mediante PR con checks verdes; verifique el origen "
              "y cierre este incidente con su post-mortem (HR-21).")
    return title, body


def audit(args: argparse.Namespace) -> int:
    event = load_json(args.event, "event")
    contract = load_json(args.contract, "contract")
    if not isinstance(event, dict) or not isinstance(contract, dict):
        fail("event and contract must be JSON objects")
    # #361: la proteccion del contrato decide ANTES del rango. Un push que CREA
    # una rama no gobernada no tiene rango que auditar; una rama protegida
    # creada de cero sigue siendo duda (HR-3), y eso lo resuelve audit_event.
    _, branch = audit_identity(event)
    default_branch = event.get("repository", {}).get("default_branch")
    if branch not in protected_branches(contract, default_branch):
        print(f"SKIP: branch '{branch}' is not protected by contract; nothing to audit")
        return 0
    repo, commits = audit_event(event)
    if commits:
        commits = compare_commits(event, repo)
    required = required_checks(contract)
    tracker = tracker_branch(contract)
    governed = governed_path_config(contract)
    findings: list[str] = []
    _, after = push_range(event)
    covered, covering = merge_coverage(repo, branch, after)
    push = PushAudit(repo, branch, event, after, tracker, required, governed, set(),
                     covered, tuple(covering), args.wait_seconds, args.poll_seconds)
    for commit in commits:
        findings += commit_findings(push, commit)
    if not findings:
        print(f"OK: {len(commits)} pushed commit(s) on '{branch}' "
              "carry a merged PR and green checks")
        return 0
    title, body = incident_report(event, branch, commits, findings)
    url = gh_open_incident(repo, title, body)
    for finding in findings:
        print(f"  VIOLATION  {finding.splitlines()[0]}")
    print(f"INCIDENT ISSUE: {url}")
    return 1


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="post-push detective vs the host contract")
    parser.add_argument("--event", required=True)
    parser.add_argument("--contract", required=True)
    parser.add_argument("--wait-seconds", type=float, default=DEFAULT_WAIT,
                        help="how long to wait for a required check the CI has not published yet")
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL,
                        help="seconds between re-reads of the required checks")
    args = parser.parse_args()
    try:
        if args.wait_seconds < 0:
            fail("--wait-seconds must be zero or more")
        if args.poll_seconds <= 0:
            fail("--poll-seconds must be positive")
        code = audit(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
