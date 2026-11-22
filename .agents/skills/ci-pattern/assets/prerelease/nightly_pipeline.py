#!/usr/bin/env python3
"""``plan``: DEPLOY o HOLD. ``candidate``: el sha de deploy. Solo lee; prefijos ``<caller> / `` solo con
``--caller``; sin ``--caller`` solo nombres sin prefijo. Cualquier lectura fallida retiene. La API de GitHub se
llama con ``runner-controls/github_rest.py`` (urllib, stdlib), no con el CLI ``gh``: el runner del consumer no lo
trae (#1691). Desde #1724 la noche despliega el sha verde mas reciente que descienda de la punta de main.
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
DEPLOY_ENVIRONMENT = "production"
MAX_CANDIDATES = 10  # Ejecuciones examinadas para elegir el sha de deploy.
#: La propia nocturna no es evidencia de su sha: construye el candidato, no su propia cabeza (#1724).
DEFAULT_EXCLUDED_WORKFLOWS = ("nightly-pipeline",)

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


def check_jobs(jobs: dict, required: tuple, pending: int) -> tuple[bool, str]:
    """``(ok, motivo)``: los jobs requeridos en verde y sin migraciones pendientes."""
    missing = [job for job in required if job not in jobs]
    if missing:
        return False, "required job not in the run: " + "; ".join(f"{job} missing" for job in missing)
    bad = [f"{job}={jobs[job]}" for job in required if jobs[job] != "success"]
    if bad:
        return False, "not all green: " + ", ".join(bad)
    if pending:
        return False, (f"{pending} pending migration(s): the nightly never applies one alone; "
                       f"apply it with a verified backup first")
    return True, f"all {len(required)} jobs green"


def decide(run_sha: str, main_sha: str, jobs: dict, required: tuple, pending: int) -> tuple[bool, str]:
    if main_sha != run_sha:
        return False, (f"the run built {run_sha[:12]} but main is now {str(main_sha)[:12]}: "
                       f"re-dispatch the pipeline on the head")
    ok, reason = check_jobs(jobs, required, pending)
    return (ok, f"{reason} on {run_sha[:12]}") if ok else (ok, reason)


def collect_jobs(jobs_list: list, caller: str | None) -> dict:
    """El resultado de cada job del consumer: un nombre repetido cuenta por su peor resultado."""
    jobs: dict = {}
    for job in jobs_list:
        if not (isinstance(job, dict) and isinstance(job.get("name"), str)):
            continue
        value = job.get("conclusion") if job.get("status") == "completed" else None
        name = job_key(job["name"], caller)
        if name is None:
            continue
        if jobs.get(name, "success") == "success":
            jobs[name] = value
    return jobs


def compare_status(repo: str, base: str, head: str) -> str:
    """El estado de la comparacion ``base...head``: ``ahead``, ``behind``, ``identical`` o ``diverged``."""
    body = api_get(f"repos/{repo}/compare/{base}...{head}")
    if not isinstance(body, dict):
        reason = f"GET repos/{repo}/compare/{base}...{head} did not return an object"
        raise ReadError(reason)
    status = body.get("status")
    if status not in ("ahead", "behind", "identical", "diverged"):
        reason = f"GET repos/{repo}/compare/{base}...{head} has no usable status"
        raise ReadError(reason)
    return status


def is_ancestor_or_equal(repo: str, ancestor: str, descendant: str) -> bool:
    """``ancestor`` es el mismo commit que ``descendant`` o desciende de el."""
    return compare_status(repo, descendant, ancestor) in ("behind", "identical")


def descends_from(repo: str, base: str, head: str) -> bool:
    """``head`` desciende de ``base``: ``base`` es su antepasado estricto."""
    return compare_status(repo, base, head) == "ahead"


def deployed_revision(repo: str, environment: str) -> str | None:
    """La revision desplegada: el deployment mas reciente con estado ``success``."""
    body = api_get(f"repos/{repo}/deployments?environment={urllib.parse.quote(environment, safe='')}"
                   f"&per_page={MAX_CANDIDATES}")
    if not isinstance(body, list):
        reason = f"GET repos/{repo}/deployments did not return a list"
        raise ReadError(reason)
    for dep in body:
        if not isinstance(dep, dict):
            continue
        number, sha = dep.get("id"), dep.get("sha")
        if not isinstance(number, int) or not isinstance(sha, str) or not sha:
            continue
        states = api_get(f"repos/{repo}/deployments/{number}/statuses?per_page=1")
        if isinstance(states, list) and states and isinstance(states[0], dict) \
                and states[0].get("state") == "success":
            return sha
    return None


def normalize_workflow(name: str) -> str:
    """El nombre de un workflow, sin ruta ni extension y en minusculas."""
    normalized = name.rsplit("/", 1)[-1].strip().lower()
    for suffix in (".yml", ".yaml"):
        if normalized.endswith(suffix):
            normalized = normalized[: -len(suffix)]
    return normalized


def evidence_runs(repo: str, branch: str = "main") -> list:
    """Las ejecuciones de la rama, la mas reciente primero."""
    body = api_get(f"repos/{repo}/actions/runs?branch={branch}&per_page={MAX_CANDIDATES}")
    runs = body.get("workflow_runs") if isinstance(body, dict) else None
    if not isinstance(runs, list):
        reason = f"GET repos/{repo}/actions/runs did not return a run list"
        raise ReadError(reason)
    return runs


def read_main_head(repo: str) -> str:
    """La cabeza de ``main``: su sha."""
    ref = api_get(f"repos/{repo}/git/ref/heads/main")
    sha = ((ref.get("object") or {}).get("sha")) if isinstance(ref, dict) else None
    if not isinstance(sha, str) or not sha:
        reason = "main has no head SHA"
        raise ReadError(reason)
    return sha


def is_green_run(repo: str, run: dict, args: argparse.Namespace) -> bool:
    """La ejecucion tiene los jobs requeridos en verde."""
    jobs = collect_jobs(api_jobs(f"repos/{repo}/actions/runs/{run['id']}/jobs?per_page=100"), args.caller)
    mapped = map_jobs(jobs, args.job_map)
    return all(mapped.get(job) == "success" for job in REQUIRED_JOBS)


def select_candidate(repo: str, tip: str, deployed: str | None, runs: list,
                     args: argparse.Namespace) -> tuple[str, str]:
    """El sha de deploy: el verde mas reciente que descienda de la punta y de la revision desplegada.

    Sin evidencia verde anterior, el candidato es la punta: la noche la prueba en verde.
    """
    seen = set()
    for run in runs:
        if not isinstance(run, dict):
            continue
        sha = run.get("head_sha")
        if not isinstance(sha, str) or not sha or sha in seen:
            continue
        seen.add(sha)
        name = run.get("path") or run.get("name")
        if normalize_workflow(name if isinstance(name, str) else "") in args.excluded:
            continue
        if run.get("status") != "completed" or run.get("conclusion") != "success":
            continue
        if not is_green_run(repo, run, args):
            continue
        if not is_ancestor_or_equal(repo, sha, tip):
            continue
        if deployed is not None and not descends_from(repo, deployed, sha):
            continue
        return sha, f"the newest green ancestor of the main head {tip[:12]}"
    return tip, f"no green ancestor of the main head {tip[:12]}: the candidate is the main head"


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
    jobs = collect_jobs(jobs_list, args.caller)
    mapped = map_jobs(jobs, args.job_map)
    if args.candidate:
        return evaluate_candidate(args, args.candidate, main_sha, mapped)
    return decide(run_sha, main_sha, mapped, REQUIRED_JOBS, args.pending_migrations)


def evaluate_candidate(args: argparse.Namespace, candidate: str, main_sha: str, jobs: dict) -> tuple[bool, str]:
    """El gate con candidato: en la historia de la punta, descendiente de la desplegada y en verde."""
    try:
        deployed = deployed_revision(args.repo, args.environment)
        if not is_ancestor_or_equal(args.repo, candidate, main_sha):
            return False, (f"the candidate {candidate[:12]} is not in the history of the main head "
                           f"{main_sha[:12]}")
        if deployed is not None and not descends_from(args.repo, deployed, candidate):
            return False, (f"the candidate {candidate[:12]} does not descend from the deployed revision "
                           f"{deployed[:12]}: refusing to roll back")
    except ReadError as exc:
        return False, (f"could not read the nightly state: {exc}. Unlock with: "
                       f"redispatch_nightly.py --repo {args.repo}")
    ok, reason = check_jobs(jobs, REQUIRED_JOBS, args.pending_migrations)
    return (ok, f"{reason} on {candidate[:12]}") if ok else (ok, reason)


def cmd_candidate(args: argparse.Namespace) -> int:
    """El sha de deploy: el verde mas reciente que descienda de la punta de ``main`` (#1724)."""
    try:
        args.job_map = parse_job_map(args.job)
    except JobMapError as exc:
        print(f"HOLD: {exc}")
        return 2
    excluded = {normalize_workflow(name) for name in DEFAULT_EXCLUDED_WORKFLOWS}
    excluded.update(normalize_workflow(name) for name in (args.exclude_workflow or ()))
    args.excluded = excluded
    try:
        if not read_profile(Path(args.profile)):
            print("HOLD: profile is not enabled in .github/prerelease.json: the nightly does not deploy")
            return 2
        tip = read_main_head(args.repo)
        deployed = deployed_revision(args.repo, args.environment)
        runs = evidence_runs(args.repo)
        candidate, why = select_candidate(args.repo, tip, deployed, runs, args)
    except (ProfileError, ReadError) as exc:
        print(f"HOLD: could not read the nightly state: {exc}. Unlock with: "
              f"redispatch_nightly.py --repo {args.repo}")
        return 2
    print(f"CANDIDATE: {candidate}")
    print(f"DEPLOYED: {deployed or 'none'}")
    print(f"WHY: {why}")
    if args.github_output:
        with Path(args.github_output).open("a", encoding="utf-8") as handle:
            handle.write(f"sha={candidate}\n")
            handle.write(f"deployed={deployed or 'none'}\n")
    return 0


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
    plan.add_argument("--candidate", default=None,
                      help=("sha de deploy elegido por el job candidate; sin el, el sha de la ejecucion"))
    plan.add_argument("--environment", default=DEPLOY_ENVIRONMENT,
                      help=("entorno del deploy, cuya revision desplegada es la base (por defecto "
                            "production)"))
    pick = sub.add_parser("candidate", help="elige el sha de deploy; solo lee")
    pick.add_argument("--repo", required=True, help="owner/name")
    pick.add_argument("--profile", required=True, help="ruta de .github/prerelease.json")
    pick.add_argument("--caller", default=None,
                      help="prefijo <caller> de workflow_call que se acepta; sin el, solo nombres sin prefijo")
    pick.add_argument("--job", action="append", default=None, metavar="REQUIRED=CONSUMER",
                      help=("lee el job requerido REQUIRED del job CONSUMER del consumer (repetible), "
                            "p. ej. --job e2e-full=e2e --job trivy=build"))
    pick.add_argument("--environment", default=DEPLOY_ENVIRONMENT,
                      help=("entorno del deploy, cuya revision desplegada es la base (por defecto "
                            "production)"))
    pick.add_argument("--exclude-workflow", action="append", default=None, metavar="NAME",
                      help=("workflow que no es evidencia del sha que construye (repetible; por defecto, "
                            "el de la propia nocturna)"))
    pick.add_argument("--github-output", default=None, help="fichero de outputs de GitHub")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.command == "candidate":
        return cmd_candidate(args)
    return cmd_plan(args)


if __name__ == "__main__":
    sys.exit(main())
