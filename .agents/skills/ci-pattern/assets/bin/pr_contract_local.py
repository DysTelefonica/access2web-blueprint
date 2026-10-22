#!/usr/bin/env python3
# ci-pattern asset — contrato de PR en local (#912)
"""Contrato de PR en local (#912): evento desde cuerpo y git, y gate en proceso.

Escribe el evento y los datos del rango (`base...HEAD`) y llama a
`assets/pr-contract/check_pr_contract.py`: veredicto, textos y exit codes son los del CI.
Sin red ni token.
"""

from __future__ import annotations

import argparse
import importlib.machinery
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "pr-contract" / "check_pr_contract.py"
DEFAULT_POLICY = ".github/pr-contract.policy.json"
REMOTE_RX = re.compile(r"github\.com[:/]([^/\s]+/[^/\s]+?)(?:\.git)?/?$")


class UsageError(RuntimeError):
    """Entrada que el comando no puede interpretar; sale con 2 y un mensaje.

    Hereda de ``RuntimeError`` porque es el contrato que espera el runner de
    git que el runner local inyecta al módulo compartido del presupuesto
    (#931): quien lo llama trata un comando fallido como origen desconocido.
    """


def _pin_output_encoding():
    """UTF-8 fijo en stdout y stderr: la salida no depende del locale (como el gate)."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def _git(root, *args):
    # git en PATH, args de control; sin shell.
    proc = subprocess.run(["git", "-C", str(root), *args],  # noqa: S603, S607
                          capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        raise UsageError(f"git {' '.join(args)} falla: {proc.stderr.strip()}")  # noqa: TRY003
    return proc.stdout


def repo_full_name_from_url(url):
    """owner/repo de una URL o de una dirección SSH de GitHub; vacío si no es de GitHub."""
    match = REMOTE_RX.search(url.strip())
    return match.group(1) if match else ""


def resolve_base(root, name):
    """La base del rango: ``origin/<name>`` si existe, si no ``<name>``."""
    for ref in (f"origin/{name}", name):
        proc = subprocess.run(  # noqa: S603
            ["git", "-C", str(root), "rev-parse", "--verify", "-q", ref],  # noqa: S607
            capture_output=True, text=True, check=False)
        if proc.returncode == 0:
            return ref
    raise UsageError(f"la base '{name}' no resuelve (ni origin/{name} ni {name}): "  # noqa: TRY003
                     "haga fetch o indique otra con --base.")


def _load_budget():
    """Módulo compartido del presupuesto (#931): fuente única de la regla A1."""
    budget_dir = str(Path(__file__).resolve().parents[1] / "budget")
    if budget_dir not in sys.path:
        sys.path.insert(0, budget_dir)
    import vendored_budget  # noqa: PLC0415

    return vendored_budget


def origins_for(root, base, head):
    """Fichero de origen del PR (#931): el módulo no ejecuta git, se lo inyectamos."""
    budget = _load_budget()

    def run_git(*args):
        return _git(root, *args)

    return budget.origins_for(base, run_git, head)


def range_data(root, base, head):
    """Textos de los ficheros del rango, con la forma que espera el gate."""
    rng = f"{base}...{head}"
    return {
        "changed": _git(root, "diff", "--name-only", rng),
        "numstat": _git(root, "diff", "--numstat", rng),
        "commits": _git(root, "log", "--format=%P%n%B%n---COMMIT-END---", f"{base}..{head}"),
        "tracked": _git(root, "ls-files"),
        "origins": origins_for(root, base, head),
    }


def build_event(body, branch, base_ref, labels, meta):
    """Evento `pull_request` con la forma que el gate lee.

    `meta` lleva number (opcional), action, full_name y default_branch."""
    pr = {"body": body,
          "labels": [{"name": name} for name in labels],
          "base": {"ref": base_ref},
          "head": {"ref": branch}}
    if meta.get("number") is not None:
        pr["number"] = meta["number"]
    return {"action": meta["action"], "pull_request": pr,
            "repository": {"full_name": meta["full_name"],
                           "default_branch": meta["default_branch"]}}


def _load_gate():
    name = "ci_pattern_pr_contract_gate"
    spec = importlib.util.spec_from_loader(
        name, importlib.machinery.SourceFileLoader(name, str(GATE)))
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def _parser():
    parser = argparse.ArgumentParser(prog="ci-pattern pr contract",
                                     description="Contrato de PR en local (#912).")
    parser.add_argument("--body", required=True, help="Fichero con el cuerpo del PR.")
    parser.add_argument("--repo", default=".", help="Repositorio git del PR.")
    parser.add_argument("--policy-file", help=f"Política del gate (por defecto {DEFAULT_POLICY}).")
    parser.add_argument("--base", default="main", help="Rama base del PR (por defecto main).")
    parser.add_argument("--head", default="HEAD", help="Revisión de la punta (por defecto HEAD).")
    parser.add_argument("--branch", help="Rama de cabecera (por defecto la rama actual).")
    parser.add_argument("--label", action="append", default=[],
                        help="Etiqueta declarada en el PR; repetible.")
    parser.add_argument("--number", type=int, help="Número del PR, si ya existe.")
    parser.add_argument("--action", default="opened", help="Acción del evento (por defecto opened).")
    parser.add_argument("--full-name", help="owner/repo (por defecto, el remoto origin).")
    parser.add_argument("--default-branch", default="main", help="Rama por defecto del repo.")
    parser.add_argument("--print-payload", action="store_true",
                        help="Imprime el evento que recibe el gate y sale 0 sin veredicto.")
    return parser


def _run(args):
    root = Path(args.repo).resolve()
    body_path = Path(args.body)
    if not body_path.is_file():
        raise UsageError(f"no se puede leer el cuerpo (--body): {body_path}")  # noqa: TRY003
    policy_path = Path(args.policy_file) if args.policy_file else root / DEFAULT_POLICY
    if not policy_path.is_file():
        raise UsageError(f"no hay política del contrato: {policy_path} (--policy-file)")  # noqa: TRY003
    branch = args.branch or _git(root, "rev-parse", "--abbrev-ref", "HEAD").strip()
    base = resolve_base(root, args.base)
    full_name = args.full_name or repo_full_name_from_url(
        _git(root, "remote", "get-url", "origin"))
    if not full_name:
        raise UsageError("no se puede leer owner/repo del remoto origin: use --full-name.")  # noqa: TRY003
    meta = {"number": args.number, "action": args.action, "full_name": full_name,
            "default_branch": args.default_branch}
    event = build_event(body_path.read_text(encoding="utf-8"), branch, args.base,
                        args.label, meta)
    if args.print_payload:
        print(json.dumps(event, indent=2, ensure_ascii=False))
        return 0
    return _gate_on_range(root, event, policy_path, base, args.head)


def _gate_on_range(root, event, policy_path, base, head):
    """Escribe el evento y los datos del rango en ficheros temporales y corre el gate."""
    texts = range_data(root, base, head)
    workdir = Path(tempfile.mkdtemp(prefix="pr-contract-local-"))
    written = []
    try:
        paths = {}
        contents = {"event": json.dumps(event), **texts}
        for name, text in contents.items():
            path = workdir / f"{name}.txt"
            path.write_text(text, encoding="utf-8")
            written.append(path)
            paths[name] = str(path)
        return _load_gate().main([
            "--event-file", paths["event"], "--policy-file", str(policy_path),
            "--repo", str(root), "--changed-files", paths["changed"],
            "--numstat-file", paths["numstat"], "--commits-file", paths["commits"],
            "--tracked-files", paths["tracked"], "--origins-file", paths["origins"]])
    finally:
        for path in written:
            path.unlink(missing_ok=True)


def main(argv=None):
    _pin_output_encoding()
    args = _parser().parse_args(argv)
    try:
        return _run(args)
    except UsageError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
