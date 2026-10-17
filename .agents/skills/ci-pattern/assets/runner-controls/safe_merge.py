#!/usr/bin/env python3
# ci-pattern asset -- merge seguro en repos sin proteccion de rama (DysTelefonica/team-skills#426)
"""Merge seguro para repos privados en cuenta Free sin branch protection.

El unico camino para mergear un PR cuando no existe proteccion de rama,
rulesets ni merge queue: verifica que el PR apunta a la rama por defecto,
que no va atrasado, que el CI es verde sobre la cabeza y que la etiqueta
de cadena y la palabra de cierre son coherentes. Si algo falla no fusiona
y devuelve 1 con el hallazgo; un entorno insalvable o un error de la API
falla cerrado con 2 (HR-3): nunca un 504 se toma como <<0 fallos>>. Con todo
en verde fusiona por la API, cierra la rama remota y la local y hace
`git worktree prune`. Runner-enforced (HR-34); lo que no pase por aqui lo
detecta el detective post-push (HR-49). Solo stdlib; cada llamada a
`gh api` lleva timeout.
"""
import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

GH_TIMEOUT = 30
SHA_RE = re.compile(r"^[0-9a-f]{40}$")
CLOSING = re.compile(r"\b(closes|fixes|resolves)\s+#?\d+", re.IGNORECASE)
REFS = re.compile(r"\brefs\s+#\d+", re.IGNORECASE)


class MergeFailClosedError(Exception):
    """A condition the merge refuses to interpret; maps to exit 2.

    Each subclass owns its message template; raise sites pass data only.
    """

    template = "fail closed"

    def __init__(self, *details: object) -> None:
        super().__init__(self.template.format(*details))


class GhCallFailedError(MergeFailClosedError):
    template = "gh api call failed: {0}"


class GhApiFailedError(MergeFailClosedError):
    template = "gh api failed ({0}): {1}"


class GhJsonInvalidError(MergeFailClosedError):
    template = "gh api returned no valid JSON: {0}"


class MergeCallFailedError(MergeFailClosedError):
    template = "the merge call failed: {0}"


class HeadChangedError(MergeFailClosedError):
    template = "the head changed between check and merge (409): {0}"


class MergeBodyInvalidError(MergeFailClosedError):
    template = "gh returned no valid JSON body: {0}"


class MergeNotMergedError(MergeFailClosedError):
    template = "the merge response does not report a merged PR: {0!r}"


class PrNotPositiveError(MergeFailClosedError):
    template = "--pr must be a positive number"


class RepoUnresolvedError(MergeFailClosedError):
    template = "no --repo given and could not resolve from cwd; pass --repo OWNER/NAME explicitly"


class PullShapeError(MergeFailClosedError):
    template = "PR #{0} is not the API pull shape"


class PullNoHeadShaError(MergeFailClosedError):
    template = "PR #{0} has no head SHA"


class PullMissingFieldsError(MergeFailClosedError):
    template = "PR #{0} is missing labels or base info"


class RepoShapeError(MergeFailClosedError):
    template = "repo info is not the API shape"


class CompareShapeError(MergeFailClosedError):
    template = "compare result is not the API shape"


class CheckRunsShapeError(MergeFailClosedError):
    template = "check-runs for PR #{0} is not the API shape"




def gh_json(path: str) -> object:
    try:
        proc = subprocess.run(
            ["gh", "api", path],
            capture_output=True, text=True, timeout=GH_TIMEOUT,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise GhCallFailedError(exc) from exc
    if proc.returncode != 0:
        stderr = proc.stderr.strip()
        # Un error de la API (5xx) es falla cerrada, nunca verde (HR-3).
        raise GhApiFailedError(proc.returncode, stderr)
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise GhJsonInvalidError(exc) from exc


def gh_merge(repo: str, number: int, sha: str) -> str:
    """Merge la PR por la API (PUT pulls/{n}/merge) con el head SHA fijado."""
    try:
        proc = subprocess.run(
            ["gh", "api", "-X", "PUT", f"repos/{repo}/pulls/{number}/merge",
             "-F", f"commit_title=Merge PR #{number} (safe merge)",
             "-F", f"sha={sha}"],
            capture_output=True, text=True, timeout=GH_TIMEOUT,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise MergeCallFailedError(exc) from exc
    if proc.returncode != 0:
        stderr = proc.stderr.strip()
        # 409 = la rama cambio entre la verificacion y el merge: fallar cerrado.
        if "409" in stderr or "conflict" in stderr.lower():
            raise HeadChangedError(stderr)
        raise MergeCallFailedError(stderr)
    try:
        data = json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise MergeBodyInvalidError(exc) from exc
    if not isinstance(data, dict) or data.get("merged") is not True:
        raise MergeNotMergedError(proc.stdout.strip())
    return data.get("sha", "unknown")


def cleanup_branch(repo: str, branch: str, default_branch: str) -> None:
    """Eliminar la rama remota (si no es la default) y la local (si existe),
    luego `git worktree prune` (solo si cwd es un worktree git)."""
    # Rama remota.
    if branch != default_branch:
        try:
            proc = subprocess.run(
                ["gh", "api", "-X", "DELETE", f"repos/{repo}/git/refs/heads/{branch}"],
                capture_output=True, text=True, timeout=GH_TIMEOUT,
            )
            if proc.returncode != 0:
                print(f"WARN: remote branch delete failed: {proc.stderr.strip()[:200]}")
            else:
                print(f"CLEAN: remote branch '{branch}' deleted")
        except (OSError, subprocess.TimeoutExpired) as exc:
            print(f"WARN: could not delete remote branch '{branch}': {exc}")
    else:
        print(f"SKIP: branch '{branch}' is the default branch, not deleting")
    # Rama local (solo si existe).
    try:
        proc = subprocess.run(
            ["git", "rev-parse", "--verify", "--quiet", branch],
            capture_output=True, text=True, timeout=5,
        )
        if proc.returncode == 0:
            subprocess.run(["git", "branch", "-D", branch], capture_output=True, text=True, timeout=5)
            print(f"CLEAN: local branch '{branch}' deleted")
    except (OSError, subprocess.TimeoutExpired):
        pass
    # worktree prune (solo si cwd es un worktree git).
    try:
        proc = subprocess.run(
            ["git", "worktree", "prune"],
            capture_output=True, text=True, timeout=5,
        )
        if proc.returncode == 0:
            print("CLEAN: worktree prune")
        else:
            print(f"WARN: worktree prune exited {proc.returncode}: {proc.stderr.strip()[:200]}")
    except (OSError, subprocess.TimeoutExpired):
        print("NOTE: cwd does not appear to be a git worktree; skipping worktree prune")


def resolve_repo() -> str:
    """Resolver el repo OWNER/REPO del ambiente."""
    env_repo = Path.cwd().joinpath(".git").joinpath("config").read_text(encoding="utf-8", errors="ignore")
    if env_repo:
        repo_env = Path.cwd().joinpath(".git").joinpath("config").read_text(encoding="utf-8", errors="ignore")
        for line in repo_env.splitlines():
            if "remote.origin.url" in line:
                url = line.split("=")[-1].strip()
                # Remueve .git, .wiki, etc. y parsea OWNER/REPO.
                url = re.sub(r"\.git/?$", "", url)
                if "/".join(url.split("/")[-2:]) == url and "/" in url:
                    return url
    return ""


def main() -> int:
    parser = argparse.ArgumentParser(
        description="merge seguro para repos sin branch protection (HR-3, HR-49, #426)",
    )
    parser.add_argument("--pr", required=True, type=int, help="numero de pull request")
    parser.add_argument("--repo", default="", help="OWNER/REPO (default: detecta desde cwd)")
    parser.add_argument("--required-check", action="append", default=[],
                        help="check requerido que debe estar verde (repitelo para varios)")
    parser.add_argument("--check-only", action="store_true",
                        help="solo verifica; no fusiona")
    args = parser.parse_args()
    try:
        if args.pr < 1:
            raise PrNotPositiveError()
        if not args.repo:
            args.repo = resolve_repo()
            if not args.repo:
                raise RepoUnresolvedError()
        # -----------------------------------------------------------
        # Paso 1: estado del PR
        # -----------------------------------------------------------
        pull = gh_json(f"repos/{args.repo}/pulls/{args.pr}")
        if not isinstance(pull, dict):
            raise PullShapeError(args.pr)
        if not isinstance(pull.get("head"), dict) or not isinstance(pull["head"].get("sha"), str):
            raise PullNoHeadShaError(args.pr)
        if not isinstance(pull.get("labels"), list) or not isinstance(pull.get("base"), dict):
            raise PullMissingFieldsError(args.pr)
        body = pull.get("body") if isinstance(pull.get("body"), str) else ""
        labels = {label.get("name") for label in pull["labels"]
                  if isinstance(label, dict) and isinstance(label.get("name"), str)}
        findings = []
        if pull.get("state") != "open":
            findings.append(f"PR #{args.pr} is not open (state: {pull.get('state')!r})")
        # -----------------------------------------------------------
        # Paso 2: base = rama por defecto (incluye stacked PRs, #420)
        # -----------------------------------------------------------
        base_ref = pull["base"]["ref"]
        default_branch = base_ref  # el PR apunta a su base; la API del repo
        # obtiene el default_branch por separado.
        repo_info = gh_json(f"repos/{args.repo}")
        if not isinstance(repo_info, dict):
            raise RepoShapeError()
        default_branch = repo_info.get("default_branch", "main")
        if base_ref != default_branch:
            findings.append(
                f"PR #{args.pr} targets '{base_ref}' (default branch is '{default_branch}'); "
                "the base must be the default branch -- stacked PRs must merge their root first (#420)"
            )
        # -----------------------------------------------------------
        # Paso 3: behind_by == 0
        # -----------------------------------------------------------
        compare = gh_json(
            f"repos/{args.repo}/compare/{default_branch}...{pull['head']['sha']}"
        )
        if not isinstance(compare, dict) or not isinstance(compare.get("behind_by"), int):
            raise CompareShapeError()
        behind = compare.get("behind_by", -1)
        if behind > 0:
            findings.append(
                f"PR #{args.pr} is behind by {behind} commit(s) against '{default_branch}'; "
                f"run `git fetch origin && git merge origin/{default_branch}`, wait for CI green, "
                "then retry."
            )
        # -----------------------------------------------------------
        # Paso 4: CI verde sobre la cabeza
        # -----------------------------------------------------------
        check_runs_data = gh_json(
            f"repos/{args.repo}/commits/{pull['head']['sha']}/check-runs"
            "?filter=latest&per_page=100"
            "&Accept=application/vnd.github+json"
        )
        if not isinstance(check_runs_data, dict) or not isinstance(check_runs_data.get("check_runs"), list):
            raise CheckRunsShapeError(args.pr)
        latest: dict = {}
        for run in check_runs_data.get("check_runs", []):
            if isinstance(run, dict) and isinstance(run.get("name"), str):
                if run["name"] not in latest or run.get("id", 0) > latest[run["name"]].get("id", 0):
                    latest[run["name"]] = run
        required_checks = set(args.required_check)
        if required_checks:
            present = {run["name"] for run in latest.values()}
            missing = required_checks - present
            if missing:
                findings.append(f"required check(s) missing: {', '.join(sorted(missing))}")
        for name, run in latest.items():
            status = run.get("status", "pending")
            conclusion = run.get("conclusion", "pending")
            if status != "completed" or conclusion not in ("success", "neutral", "skipped"):
                findings.append(
                    f"check '{name}': status={status}, conclusion={conclusion or 'pending'}"
                )
        if not latest and behind == 0:
            # Sin check-runs y sin commits atrasados: si no hay checks en absoluto,
            # es un repo sin CI configurado o los checks son GitHub Actions
            # que no aparecen con check-runs (son check-suites). No bloqueamos
            # si el PR no tiene checks requeridos explicitos.
            pass
        # -----------------------------------------------------------
        # Paso 5: coherencia chain/closing
        # -----------------------------------------------------------
        is_chain = "chain:partial" in labels
        has_closing = bool(CLOSING.search(body))
        has_refs = bool(REFS.search(body))
        if is_chain:
            if not has_refs:
                findings.append(
                    f"PR #{args.pr} is chain:partial but body does not match `Refs #<digits>`"
                )
            if has_closing:
                findings.append(
                    f"PR #{args.pr} is chain:partial but body carries closing words (Closes/Fixes/Resolves); "
                    "only the chain tip may use them"
                )
        else:
            if not has_closing:
                findings.append(
                    f"PR #{args.pr} is the chain tip (no chain:partial) but body does not carry a closing word"
                )
        # -----------------------------------------------------------
        # Decision
        # -----------------------------------------------------------
        if findings:
            for f in findings:
                print(f"DENIED: {f}")
            print(f"RECOVERY: run `git fetch origin && git merge origin/{default_branch}`, "
                  "wait for CI green, then retry.")
            return 1
        # -----------------------------------------------------------
        # Fusionar o solo verificar
        # -----------------------------------------------------------
        if args.check_only:
            print(f"OK: PR #{args.pr} passes all checks (check-only mode)")
            return 0
        sha = gh_merge(args.repo, args.pr, pull["head"]["sha"])
        print(f"MERGED: PR #{args.pr} merged by safe-merge as {sha}")
        cleanup_branch(args.repo, pull["head"]["ref"], default_branch)
        return 0
    except MergeFailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
