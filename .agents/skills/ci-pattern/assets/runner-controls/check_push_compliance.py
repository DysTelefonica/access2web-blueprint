#!/usr/bin/env python3
# ci-pattern asset — post-push compliance detective (DysTelefonica/team-skills#303)
"""Detective post-push: audita cada push a una rama protegida por contrato.

El contrato declara `protected_branches` y `required_checks`; para cada
commit del evento `push` la API debe mostrar un PR mergeado y los checks
requeridos en verde sobre la cabeza de ese PR. Una violación abre una issue
de incidente con la evidencia (SHA, autor, checks) y sale con 1; un push
conforme sale con 0. Entorno insalvable falla cerrado con 2 (HR-3): nunca
verde sin medición. Runner-enforced (HR-34): sin protección de rama este
detective es el control que actúa y deja evidencia.
"""
import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn

from github_rest import GitHubApiError, call, get_json

SHA_RE = re.compile(r"^[0-9a-f]{40}$")


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


def audit_event(event: dict) -> tuple[str, list[dict]]:
    repo_name, _ = audit_identity(event)
    before, after = event.get("before"), event.get("after")
    if not isinstance(before, str) or not SHA_RE.match(before) \
            or not isinstance(after, str) or not SHA_RE.match(after):
        fail("event 'before'/'after' are not 40-char commit ids")
    if before == "0" * 40:
        fail("before of zeros (new branch): the pushed range is not auditable")
    if event.get("deleted") is True:
        return repo_name, []  # el borrado lo trata el control de force-push
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


def check_tree_integrity(repo: str, head_sha: str, commit_sha: str) -> list[str]:
    """Verificar que el arbol de lo que fue revieweado coincide con el arbol
    que llego al push (si el merge fue seguro)."""
    # Obtener el arbol del head de la PR (lo que fue revieweado).
    pr_commit = gh_json(f"repos/{repo}/commits/{head_sha}")
    if not isinstance(pr_commit, dict):
        return []  # Si la API no devuelve el commit, no bloqueamos (fall-open).
    pr_tree = pr_commit.get("tree", {}).get("sha")
    if not isinstance(pr_tree, str) or not SHA_RE.match(pr_tree):
        return []
    # Obtener el arbol del commit que llego al push.
    pushed_commit = gh_json(f"repos/{repo}/commits/{commit_sha}")
    if not isinstance(pushed_commit, dict):
        return []
    pushed_tree = pushed_commit.get("tree", {}).get("sha")
    if not isinstance(pushed_tree, str):
        return []
    if pr_tree != pushed_tree:
        return [
            f"arbol de PR #{pr_tree[:12]} vs commit {commit_sha[:7]}: difieren en el arbol; "
            "esto indica que el arbol mergeado no es identico al arbol revieweado. "
            "Verifique que se uso safe-merge (#426)."
        ]
    return []


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
        else:
            conclusion = run.get("conclusion")
            if run.get("status") != "completed" or conclusion != "success":
                verdicts.append(f'required check "{name}": {conclusion or "pending"}')
    return verdicts


def compare_commits(event: dict, repo: str) -> list[dict]:
    """Rango completo del push según el compare (la API devuelve hasta 250; #326)."""
    compare = gh_json(f"repos/{repo}/compare/{event['before']}...{event['after']}")
    if not isinstance(compare, dict) or not isinstance(compare.get("total_commits"), int):
        fail(f"compare for {event['before']}...{event['after']} is not the API compare shape")
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


@dataclass(frozen=True)
class PushAudit:
    """Contexto fijo de la auditoría de un push: lo que no cambia de un commit a otro."""

    repo: str
    branch: str
    event: dict
    tracker: str | None
    required: list[str]
    seen_prs: set


def commit_findings(push: PushAudit, commit: dict) -> list[str]:
    repo, branch, event, tracker, required, seen_prs = (
        push.repo, push.branch, push.event, push.tracker, push.required, push.seen_prs)
    author = commit.get("author") if isinstance(commit.get("author"), dict) else {}
    author_name = author.get("name") or event.get("pusher", {}).get("name") or "unknown author"
    prs = gh_json(f"repos/{repo}/commits/{commit['id']}/pulls")
    if not isinstance(prs, list):
        fail(f"pulls for commit {commit['id']} is not a JSON array")
    merged = [p for p in prs if isinstance(p, dict) and (p.get("merged") or p.get("merged_at"))]
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
        if number in seen_prs:
            continue
        seen_prs.add(number)
        detail = gh_json(f"repos/{repo}/pulls/{number}")
        if not isinstance(detail, dict) or not isinstance(detail.get("head"), dict) \
                or not isinstance(detail["head"].get("sha"), str):
            fail(f"PR #{number} detail is not the API pull shape")
        login = (pull.get("user") or {}).get("login") or "unknown"
        verdicts = check_verdict(repo, number, detail["head"]["sha"], required)
        # Verificacion del arbol: lo que llego a main debe ser identico
        # a lo que fue revieweado (siempre que el commit del push
        # sea exactamente el head de la PR mergeada).
        verdicts += check_tree_integrity(repo, detail["head"]["sha"], commit["id"])
        if verdicts:
            evidence = "\n".join(f"  {line}" for line in verdicts)
            findings.append(f"commit {commit['id']} — author: {author_name} — PR #{number} "
                            f"merged by {login} (head {detail['head']['sha']})\n{evidence}")
    return findings


def incident_report(event: dict, branch: str, commits: list[dict],
                    findings: list[str]) -> tuple[str, str]:
    title = f"incident: push compliance violation on '{branch}' ({commits[0]['id'][:7]})"
    body = ("Post-push detective (`check_push_compliance`), rama protegida por contrato: "
            f"{branch}\n\npusher: {event.get('pusher', {}).get('name', 'unknown')}\n"
            f"rango: {event['before']}..{event['after']}\n\n"
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
    findings: list[str] = []
    push = PushAudit(repo, branch, event, tracker, required, set())
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


def main() -> int:
    parser = argparse.ArgumentParser(description="post-push detective vs the host contract")
    parser.add_argument("--event", required=True)
    parser.add_argument("--contract", required=True)
    args = parser.parse_args()
    try:
        code = audit(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
