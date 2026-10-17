#!/usr/bin/env python3
# ci-pattern asset — safe-cleanup helper, core (DysTelefonica/team-skills#436)
"""Nucleo de safe-cleanup: borra rama local y worktree solo si la
informacion esta preservada en otra parte. Soporte de bundle y
CLI anadidos en el tramo (b) de la cadena.

Companero de ``safe_merge.py`` (#426) para repos sin branch
protection: el merge borra la rama remota; este helper borra la
rama local y el worktree SOLO si la punta sigue en el remoto, es
ancestro de la rama activa o coincide con la cabeza de una PR
mergeada.

Invariantes (HR-3, HR-13, #436):
- rc=0 -> limpio (borrado OK o ya estaba limpio).
- rc=3 -> rehusado: el motivo nombra la causa real y como resolverlo.
- rc=2 -> error de entorno (no es repo, no hay git, etc.); falla cerrado.
- NUNCA invoca ``git push origin --delete`` ni equivalente.
- Con ``bundle_dir``, antes de negarse crea y verifica un bundle
  (incluido el camino de rechazo: el operador se queda con el
  respaldo aunque no se borre).

Uso (desde Python):

    from safe_cleanup import cleanup
    result = cleanup("/path/to/repo", "feat/x", worktree="/path/to/wt",
                      pr_head="abcdef...", bundle_dir="/backups")
    sys.exit(result.exit_code)

CLI equivalente:

    python3 safe_cleanup.py --repo-dir /path/to/repo --branch feat/x \\
        [--worktree /path/to/wt] [--pr-head abcdef...] [--bundle-dir /backups]
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

GIT_TIMEOUT = 15


class EnvironmentFailError(Exception):
    """Error de entorno: rc=2, fail-closed.

    Each subclass owns its message template; raise sites pass data only.
    """

    template = "environment failure"

    def __init__(self, *details: object) -> None:
        super().__init__(self.template.format(*details))


class GitCallFailedError(EnvironmentFailError):
    template = "git {0} failed: {1}"


class RepoDirMissingError(EnvironmentFailError):
    template = "repo_dir does not exist: {0}"


class RepoDirNotGitError(EnvironmentFailError):
    template = "repo_dir is not a git repo: {0}"


class GitStatusFailedError(EnvironmentFailError):
    template = "git status failed: {0}"


class WorktreeRemoveFailedError(EnvironmentFailError):
    template = "git worktree remove failed: {0}"


class BranchDeleteFailedError(EnvironmentFailError):
    template = "git branch -D {0} failed: {1}"




@dataclass
class Result:
    """Resultado de ``cleanup``. ``ok`` es True si se borro o ya estaba
    limpio; False si se rehuso. ``exit_code`` mapea al shell (0/3/2)."""

    ok: bool
    reason: str
    exit_code: int
    bundled_to: Path | None = None
    deleted: list[str] = field(default_factory=list)


def _run_git(args: list[str], cwd: Path) -> subprocess.CompletedProcess:
    """Ejecuta un subcomando git con timeout. Falla cerrado en OSError."""
    try:
        return subprocess.run(
            ["git", *args],
            capture_output=True, text=True, cwd=cwd,
            timeout=GIT_TIMEOUT,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise GitCallFailedError(args[0] if args else "", exc) from exc


def _resolve_repo(repo_dir: str | Path) -> Path:
    p = Path(repo_dir).resolve()
    if not p.is_dir():
        raise RepoDirMissingError(repo_dir)
    if not (p / ".git").exists():
        raise RepoDirNotGitError(repo_dir)
    return p


def _resolve_worktree(worktree: str | Path | None) -> Path | None:
    if worktree is None:
        return None
    p = Path(worktree).resolve()
    if not p.is_dir():
        return None
    return p


def _is_worktree_clean(wt: Path) -> tuple[bool, str]:
    """Comprueba ``git status --porcelain`` en el worktree."""
    proc = _run_git(["status", "--porcelain"], cwd=wt)
    if proc.returncode != 0:
        raise GitStatusFailedError(proc.stderr.strip()[:200])
    if proc.stdout.strip():
        sample = "\n".join(proc.stdout.strip().splitlines()[:5])
        return False, sample
    return True, ""


def _processes_in_path(path: Path) -> list[int]:
    """Devuelve los PIDs cuyo cwd (vía /proc/<pid>/cwd) esta dentro de path."""
    procd = Path("/proc")
    if not procd.is_dir():
        return []
    pids: list[int] = []
    target = str(path)
    for entry in procd.iterdir():
        if not entry.name.isdigit():
            continue
        try:
            cwd = os.readlink(entry / "cwd")
        except (OSError, PermissionError):
            continue
        if cwd == target or cwd.startswith(target + os.sep):
            pids.append(int(entry.name))
    return pids


def _rev(repo: Path, ref: str) -> str | None:
    """Resuelve un ref a SHA completo. None si no existe."""
    proc = _run_git(["rev-parse", "--verify", ref], cwd=repo)
    if proc.returncode != 0:
        return None
    return proc.stdout.strip()


def _active_branch(repo: Path) -> str:
    """Determina la rama activa del repo: origin/HEAD si existe, si no
    'main'. La rama activa es contra la que se mide el criterio de
    ancestro ('la punta es ancestro de origin/<rama activa>')."""
    proc = _run_git(["symbolic-ref", "--quiet", "refs/remotes/origin/HEAD"],
                    cwd=repo)
    if proc.returncode == 0 and proc.stdout.strip():
        ref = proc.stdout.strip()
        return ref.rsplit("/", 1)[-1]
    proc = _run_git(["symbolic-ref", "--quiet", "HEAD"], cwd=repo)
    if proc.returncode == 0 and proc.stdout.strip():
        return proc.stdout.strip().rsplit("/", 1)[-1]
    return "main"


def _is_preserved(
    repo: Path, branch: str, pr_head: str | None
) -> tuple[bool, str]:
    """Comprueba los tres criterios del OR de preservacion.

    Devuelve (preservado, motivo). El motivo describe el camino elegido
    cuando preservado=True; cuando es False, describe por que no."""
    tip = _rev(repo, branch)
    if tip is None:
        return False, f"local branch '{branch}' does not exist"
    if pr_head is not None and tip == pr_head:
        return True, "branch tip matches pr_head"
    remote_tip = _rev(repo, f"origin/{branch}")
    if remote_tip is not None and tip == remote_tip:
        return True, f"branch tip matches origin/{branch}"
    active = _active_branch(repo)
    proc = _run_git(
        ["merge-base", "--is-ancestor", tip, f"origin/{active}"], cwd=repo
    )
    if proc.returncode == 0:
        return True, f"branch tip is ancestor of origin/{active}"
    return False, (
        f"branch '{branch}' is not preserved: tip != origin/{branch}, "
        f"!= pr_head, and not an ancestor of origin/{active}; "
        f"push first, pass --pr-head, or escalate before retrying"
    )


def _remove_worktree(repo: Path, wt: Path) -> None:
    """``git worktree remove`` + ``git worktree prune``. Falla cerrado."""
    proc = _run_git(["worktree", "remove", "--force", str(wt)], cwd=repo)
    if proc.returncode != 0:
        raise WorktreeRemoveFailedError(proc.stderr.strip()[:200])
    _run_git(["worktree", "prune"], cwd=repo)


def _delete_local_branch(repo: Path, branch: str) -> None:
    """Borra la rama local SOLO si existe. Nunca la rama remota.

    Si la rama esta checked out en el repo principal (HEAD == branch),
    salimos primero a la rama activa para que ``git branch -D`` pueda
    operar. Nunca tocamos la rama remota."""
    proc = _run_git(["rev-parse", "--verify", "--quiet", branch], cwd=repo)
    if proc.returncode != 0:
        return
    head_branch = _run_git(
        ["symbolic-ref", "--quiet", "HEAD"], cwd=repo
    ).stdout.strip()
    if head_branch == f"refs/heads/{branch}":
        active = _active_branch(repo)
        target = active if active != branch else "main"
        _run_git(["checkout", "-q", target], cwd=repo)
    delete = _run_git(["branch", "-D", branch], cwd=repo)
    if delete.returncode != 0:
        raise BranchDeleteFailedError(branch, delete.stderr.strip()[:200])


def _create_bundle(repo: Path, branch: str, bundle_dir: Path) -> Path | None:
    """Crea un bundle de la rama y lo verifica. None si no se pudo.

    El bundle se nombra <sanitized-branch>.bundle dentro de bundle_dir.
    Falla cerrado: si la creacion o la verificacion fallan, no devolvemos
    una ruta utilizable (el caller trata como si no hubiera bundle)."""
    bundle_dir.mkdir(parents=True, exist_ok=True)
    safe_name = branch.replace("/", "_").replace("\\", "_")
    bundle_path = bundle_dir / f"{safe_name}.bundle"
    proc = _run_git(
        ["bundle", "create", str(bundle_path), branch], cwd=repo
    )
    if proc.returncode != 0:
        return None
    verify = _run_git(["bundle", "verify", str(bundle_path)], cwd=repo)
    if verify.returncode != 0:
        try:
            bundle_path.unlink()
        except OSError:
            pass
        return None
    return bundle_path


def cleanup(
    repo_dir: str | Path,
    branch: str,
    worktree: str | Path | None = None,
    pr_head: str | None = None,
    bundle_dir: str | Path | None = None,
) -> Result:
    """Borra el worktree y la rama local SOLO si la informacion esta
    preservada. Ver contrato en el docstring del modulo."""
    try:
        repo = _resolve_repo(repo_dir)
        wt = _resolve_worktree(worktree)
        bundle_path = (
            Path(bundle_dir).resolve() if bundle_dir is not None else None
        )
    except EnvironmentFailError as exc:
        return Result(ok=False, reason=str(exc), exit_code=2)

    # Pre-chequeos: con bundle_dir, el bundle se crea ANTES de negarse
    # por cualquier motivo, para que el operador tenga respaldo.
    bundle_written: Path | None = None
    if bundle_path is not None:
        try:
            bundle_written = _create_bundle(repo, branch, bundle_path)
        except EnvironmentFailError as exc:
            return Result(ok=False, reason=str(exc), exit_code=2)
        if bundle_written is None:
            return Result(
                ok=False,
                reason=(
                    f"could not create or verify bundle in '{bundle_path}' "
                    "before cleanup; refusing to proceed"
                ),
                exit_code=2,
                bundled_to=None,
            )

    if wt is not None:
        try:
            clean, dirty_sample = _is_worktree_clean(wt)
        except EnvironmentFailError as exc:
            return Result(
                ok=False, reason=str(exc), exit_code=2,
                bundled_to=bundle_written,
            )
        if not clean:
            return Result(
                ok=False,
                reason=(
                    f"worktree '{wt}' is dirty (uncommitted changes); "
                    f"commit, stash or discard before cleanup. Sample:\n"
                    f"{dirty_sample}"
                ),
                exit_code=3,
                bundled_to=bundle_written,
            )
        try:
            pids = _processes_in_path(wt)
        except EnvironmentFailError as exc:
            return Result(
                ok=False, reason=str(exc), exit_code=2,
                bundled_to=bundle_written,
            )
        if pids:
            sample = ", ".join(str(p) for p in pids[:5])
            return Result(
                ok=False,
                reason=(
                    f"worktree '{wt}' has active processes (PIDs: {sample}); "
                    "stop them or run cleanup from outside the worktree"
                ),
                exit_code=3,
                bundled_to=bundle_written,
            )

    try:
        preserved, why = _is_preserved(repo, branch, pr_head)
    except EnvironmentFailError as exc:
        return Result(
            ok=False, reason=str(exc), exit_code=2,
            bundled_to=bundle_written,
        )
    if not preserved:
        return Result(
            ok=False, reason=why, exit_code=3,
            bundled_to=bundle_written,
        )

    deleted: list[str] = []
    try:
        if wt is not None:
            _remove_worktree(repo, wt)
            deleted.append(f"worktree:{wt}")
        _delete_local_branch(repo, branch)
        deleted.append(f"branch:{branch}")
    except EnvironmentFailError as exc:
        return Result(
            ok=False, reason=str(exc), exit_code=2,
            bundled_to=bundle_written, deleted=deleted,
        )
    return Result(
        ok=True,
        reason=f"cleaned ({why})",
        exit_code=0,
        bundled_to=bundle_written,
        deleted=deleted,
    )


def _build_parser() -> argparse.ArgumentParser:  # noqa: D401
    """Construye el parser de linea de comandos para el CLI."""
    parser = argparse.ArgumentParser(
        description=(
            "safe-cleanup: borra worktree y rama local solo si la "
            "informacion esta preservada (DysTelefonica/team-skills#436)"
        ),
    )
    parser.add_argument("--repo-dir", required=True,
                        help="path al repo git")
    parser.add_argument("--branch", required=True,
                        help="rama local a borrar (la remota se conserva)")
    parser.add_argument("--worktree", default=None,
                        help="path al worktree a borrar (opcional)")
    parser.add_argument("--pr-head", default=None,
                        help="SHA del head de la PR mergeada (opcional)")
    parser.add_argument("--bundle-dir", default=None,
                        help="directorio donde escribir el bundle de respaldo "
                             "antes de cualquier borrado (opcional)")
    return parser


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    parser = _build_parser()
    args = parser.parse_args()
    result = cleanup(
        args.repo_dir,
        args.branch,
        worktree=args.worktree,
        pr_head=args.pr_head,
        bundle_dir=args.bundle_dir,
    )
    if result.bundled_to is not None:
        print(f"BUNDLE: {result.bundled_to}")
    for item in result.deleted:
        print(f"CLEAN: {item}")
    if result.ok:
        print(f"OK: {result.reason}")
    else:
        if result.exit_code == 2:
            print(f"fail-closed: {result.reason}", file=sys.stderr)
        else:
            print(f"REFUSED: {result.reason}")
    return result.exit_code


if __name__ == "__main__":
    sys.exit(main())
