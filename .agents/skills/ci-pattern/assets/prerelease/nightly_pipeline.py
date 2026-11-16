#!/usr/bin/env python3
"""``plan``: DEPLOY o HOLD. Solo lee; prefijos ``<caller> / `` solo con ``--caller``; sin ``--caller`` solo
nombres sin prefijo. Cualquier lectura fallida retiene. La API de GitHub se llama con
``runner-controls/github_rest.py`` (urllib, stdlib), no con el CLI ``gh``: el runner del consumer no lo trae (#1691).
"""
import argparse
import importlib.util
import json
import sys
import urllib.parse
from pathlib import Path

REQUIRED_JOBS = ("build", "e2e-full", "integration", "upgrade-path", "trivy", "security-deep")
MAX_PAGES = 50
LINK_PARTS = 2  # <url>; rel="..."

#: El cliente REST del layout vendorizado; se carga la primera vez que hace falta.
_REST = None


class ProfileError(Exception):
    pass


class ReadError(Exception):
    pass


class JobMapError(Exception):
    pass


def parse_job_map(values: list[str] | None) -> dict[str, str]:
    """``--job <requerido>=<nombre del consumer>`` (#1699): un mapeo por job requerido."""
    mapping: dict[str, str] = {}
    for value in values or ():
        required, sep, consumer = value.partition("=")
        if not sep or not required or not consumer:
            reason = f"--job {value}: expected <required>=<consumer job>"
            raise JobMapError(reason)
        if required not in REQUIRED_JOBS:
            reason = f"--job {value}: unknown required job {required!r}"
            raise JobMapError(reason)
        if required in mapping:
            reason = f"--job {value}: required job {required!r} is already mapped"
            raise JobMapError(reason)
        mapping[required] = consumer
    return mapping


def map_jobs(jobs: dict, job_map: dict[str, str]) -> dict:
    """Los resultados por nombre requerido: el job X se lee del job del consumer ``job_map.get(X, X)``."""
    if not job_map:
        return jobs
    return {job: jobs[job_map.get(job, job)] for job in REQUIRED_JOBS if job_map.get(job, job) in jobs}


def _pin_output_encoding() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def read_profile(path: Path) -> bool:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        reason = f"profile {path} could not be read: {exc}"
        raise ProfileError(reason) from exc
    if not isinstance(data, dict):
        reason = f"profile {path} is not a JSON object"
        raise ProfileError(reason)
    version, enabled = data.get("schema_version"), data.get("enabled")
    if type(version) is not int or version != 1:
        reason = f"profile {path}: schema_version must be 1, got {version!r}"
        raise ProfileError(reason)
    if not isinstance(enabled, bool):
        reason = f"profile {path}: enabled must be true or false, got {enabled!r}"
        raise ProfileError(reason)
    return enabled


def decide(run_sha: str, main_sha: str, jobs: dict, required: tuple, pending: int) -> tuple[bool, str]:
    if main_sha != run_sha:
        return False, (f"the run built {run_sha[:12]} but main is now {str(main_sha)[:12]}: "
                       f"re-dispatch the pipeline on the head")
    missing = [job for job in required if job not in jobs]
    if missing:
        return False, "required job not in the run: " + "; ".join(f"{job} missing" for job in missing)
    bad = [f"{job}={jobs[job]}" for job in required if jobs[job] != "success"]
    if bad:
        return False, "not all green: " + ", ".join(bad)
    if pending:
        return False, (f"{pending} pending migration(s): the nightly never applies one alone; "
                       f"apply it with a verified backup first")
    return True, f"all {len(required)} jobs green on {run_sha[:12]}"


def _load_github_rest():
    """Carga ``runner-controls/github_rest.py`` del layout vendorizado (#1691).

    El asset se propaga junto a sus hermanos, no como paquete: el modulo se
    carga por ruta, como hacen ``issue-spec/check_issue_spec.py`` y
    ``pr-contract/check_pr_contract.py``.
    """
    path = Path(__file__).resolve().parents[1] / "runner-controls" / "github_rest.py"
    if not path.is_file():
        reason = f"{path} is missing"
        raise RuntimeError(reason)
    spec = importlib.util.spec_from_file_location("nightly_pipeline_github_rest", path)
    if spec is None or spec.loader is None:
        reason = f"{path} could not be loaded"
        raise RuntimeError(reason)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _rest():
    """El cliente REST; si el asset vendorizado no esta, se falla cerrado."""
    global _REST
    if _REST is None:
        try:
            _REST = _load_github_rest()
        except (OSError, RuntimeError, ImportError) as exc:
            reason = f"the GitHub REST client could not be loaded: {exc}"
            raise ReadError(reason) from exc
    return _REST


def _next_path(link: str) -> str | None:
    """Ruta y query de ``rel="next"``; el host lo fija ``GITHUB_API_URL``, nunca la cabecera."""
    for part in link.split(","):
        segments = part.strip().split(";")
        if len(segments) == LINK_PARTS and segments[1].strip() == 'rel="next"':
            parts = urllib.parse.urlsplit(segments[0].strip().strip("<>"))
            path = parts.path.lstrip("/")
            return f"{path}?{parts.query}" if parts.query else path
    return None


def api_get(path: str) -> object:
    """GET de un documento JSON; un fallo de lectura levanta ReadError."""
    rest = _rest()
    try:
        return rest.get_json(path)
    except rest.GitHubApiError as exc:
        reason = f"GET {path}: {exc}"
        raise ReadError(reason) from exc


def api_jobs(path: str) -> list:
    """Todos los jobs de todas las paginas: se sigue la cabecera ``Link``."""
    rest = _rest()
    jobs: list = []
    target = path
    for _ in range(MAX_PAGES):
        try:
            text, link = rest.call_con_reintento("GET", target)
        except rest.GitHubApiError as exc:
            reason = f"GET {target}: {exc}"
            raise ReadError(reason) from exc
        try:
            page = json.loads(text)
        except ValueError as exc:
            reason = f"GET {target} did not return JSON"
            raise ReadError(reason) from exc
        page_jobs = page.get("jobs") if isinstance(page, dict) else None
        if not isinstance(page_jobs, list):
            reason = f"GET {target}: a page has no jobs list"
            raise ReadError(reason)
        jobs.extend(page_jobs)
        nxt = _next_path(link)
        if nxt is None:
            return jobs
        target = nxt
    reason = f"GET {path} exceeds {MAX_PAGES} pages; refusing to continue"
    raise ReadError(reason)


def job_key(name: str, caller: str | None) -> str | None:
    prefix, sep, bare = name.rpartition(" / ")
    if not sep:
        return name
    return bare if caller is not None and prefix == caller else None


def evaluate(args: argparse.Namespace) -> tuple[bool, str]:
    try:
        if not read_profile(Path(args.profile)):
            return False, "profile is not enabled in .github/prerelease.json: the nightly does not deploy"
        run = api_get(f"repos/{args.repo}/actions/runs/{args.run}")
        jobs_list = api_jobs(f"repos/{args.repo}/actions/runs/{args.run}/jobs?per_page=100")
        main_ref = api_get(f"repos/{args.repo}/git/ref/heads/main")
    except (ProfileError, ReadError) as exc:
        return False, (f"could not read the nightly state: {exc}. Unlock with: "
                       f"redispatch_nightly.py --repo {args.repo}")
    run_sha = run.get("head_sha") if isinstance(run, dict) else None
    main_sha = ((main_ref.get("object") or {}).get("sha")) if isinstance(main_ref, dict) else None
    if not isinstance(run_sha, str) or not run_sha:
        return False, "the run has no head SHA: hold"
    if not isinstance(main_sha, str) or not main_sha:
        return False, "main has no head SHA: hold"
    jobs: dict = {}
    for job in jobs_list:
        if not (isinstance(job, dict) and isinstance(job.get("name"), str)):
            continue
        value = job.get("conclusion") if job.get("status") == "completed" else None
        name = job_key(job["name"], args.caller)
        if name is None:
            continue
        if jobs.get(name, "success") == "success":
            jobs[name] = value
    return decide(run_sha, main_sha, map_jobs(jobs, args.job_map), REQUIRED_JOBS, args.pending_migrations)


def cmd_plan(args: argparse.Namespace) -> int:
    if args.pending_migrations < 0:
        print("HOLD: --pending-migrations must be 0 or more")
        return 2
    try:
        args.job_map = parse_job_map(args.job)
    except JobMapError as exc:
        print(f"HOLD: {exc}")
        return 2
    deploy, reason = evaluate(args)
    print(f"{'DEPLOY' if deploy else 'HOLD'}: {reason}")
    return 0 if deploy else 2


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Decision de la canalizacion nocturna pre-release (#1627, slice 3).")
    sub = parser.add_subparsers(dest="command", required=True)
    plan = sub.add_parser("plan", help="decide DEPLOY o HOLD; solo lee")
    plan.add_argument("--repo", required=True, help="owner/name")
    plan.add_argument("--run", required=True, type=int, help="id de la ejecucion nocturna")
    plan.add_argument("--profile", required=True, help="ruta de .github/prerelease.json")
    plan.add_argument("--caller", default=None,
                      help="prefijo <caller> de workflow_call que se acepta; sin el, solo nombres sin prefijo")
    plan.add_argument("--pending-migrations", required=True, type=int,
                      help="migraciones pendientes contadas por el job build")
    plan.add_argument("--job", action="append", default=None, metavar="REQUIRED=CONSUMER",
                      help=("lee el job requerido REQUIRED del job CONSUMER del consumer (repetible), "
                            "p. ej. --job e2e-full=e2e --job trivy=build"))
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    return cmd_plan(args)


if __name__ == "__main__":
    sys.exit(main())
