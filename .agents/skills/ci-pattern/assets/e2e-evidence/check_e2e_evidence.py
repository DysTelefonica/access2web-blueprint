#!/usr/bin/env python3
# ci-pattern asset — e2e-evidence gate (DysTelefonica/team-skills#425)
"""Gate de evidencia e2e para PRs con cambio de UI.

Regla: toda PR que toque una superficie de UI anade o actualiza al
menos un flujo e2e de navegador, que entra por la navegacion real
del consumer. La exencion es explicita y esta justificada en el
cuerpo de la PR (seccion dedicada).

Parametros del consumer:

- ``--ui-glob <patron>``: glob que matchea los paths de UI del consumer
  (default: ``app/**``, ``web/**``, ``src/**``, ``pages/**``,
  ``components/**``, ``views/**``, ``templates/**``). Repetirlo reemplaza
  la lista default.
- ``--non-ui-prefix <prefix>``: prefix de paths que NO son UI
  (default: ``.atl/,.codegraph/,.github/,.pi/,docs/,git-hooks/,migration/,odd/,openspec/,scripts/,skills/,tests/,.dockerignore,.env.example,.gitignore,AGENTS.md,CHANGELOG.md,CODEOWNERS,CONTRIBUTING.md,README.md,SECURITY.md,pyproject.toml,uv.lock,Dockerfile,docker-compose.yml,Makefile``). Repetirlo reemplaza la lista default.
- ``--e2e-dir <path>``: directorio de flujos e2e relativo a la raiz
  (default: ``tests/e2e_ci``).
- ``--e2e-pattern <glob>``: patron de ficheros de flujo e2e (default:
  ``*.spec.{ts,js,py}``); admite UN grupo ``{a,b,c}`` que se expande.
- ``--exemption-section <string>``: seccion del cuerpo de la PR donde
  se documenta la exencion (default: ``## Exencion e2e``).

Exit codes:

- 0: limpio (sin UI, o UI con e2e, o UI con exencion justificada).
- 1: hallazgos (UI sin e2e y sin exencion).
- 2: entrada invalida (PR body ausente, globs malformados, git falla).
"""
from __future__ import annotations

import argparse
import fnmatch
import re
import subprocess
import sys
from pathlib import Path

DEFAULT_UI_GLOBS = ["app/**", "web/**", "src/**", "pages/**",
                    "components/**", "views/**", "templates/**"]

DEFAULT_NON_UI = [
    ".atl/", ".codegraph/", ".github/", ".pi/", "docs/", "git-hooks/",
    "migration/", "odd/", "openspec/", "scripts/", "skills/", "tests/",
    ".dockerignore", ".env.example", ".gitignore", "AGENTS.md",
    "CHANGELOG.md", "CODEOWNERS", "CONTRIBUTING.md", "README.md",
    "SECURITY.md", "pyproject.toml", "uv.lock", "Dockerfile",
    "docker-compose.yml", "Makefile",
]


def run_git(repo: Path, *args: str) -> str:
    proc = subprocess.run(
        ["git", "-C", str(repo), *args],
        capture_output=True, text=True, check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"git {' '.join(args)} fallo (rc={proc.returncode}): {proc.stderr.strip()}"
        )
    return proc.stdout


def list_changed_paths(repo: Path, base: str, head: str) -> list[str]:
    """Lista los paths cambiados entre base y head."""
    out = run_git(repo, "diff", "--name-only", f"{base}..{head}")
    return [line.strip() for line in out.splitlines() if line.strip()]


def expand_pattern(pattern: str) -> list[str]:
    """Expande un patron con un grupo ``{a,b,c}`` en la lista de patrones
    concretos que fnmatch si entiende (fnmatch no expande llamas)."""
    m = re.search(r"\{([^{}]*)\}", pattern)
    if not m:
        return [pattern]
    head, tail = pattern[:m.start()], pattern[m.end():]
    expanded = []
    for alt in m.group(1).split(","):
        expanded.extend(expand_pattern(head + alt + tail))
    return expanded


def list_added_e2e_paths(repo: Path, base: str, head: str,
                        e2e_dir: str, pattern: str) -> list[str]:
    """Lista los paths de flujos e2e anadidos o modificados entre base y
    head (la invariante admite actualizar un flujo existente)."""
    out = run_git(repo, "diff", "--name-only", "--diff-filter=AM",
                  f"{base}..{head}", "--", e2e_dir)
    patterns = expand_pattern(pattern)
    added = []
    for line in out.splitlines():
        if not line.strip():
            continue
        if any(fnmatch.fnmatch(Path(line).name, p) for p in patterns):
            added.append(line.strip())
    return added


def is_ui_path(path: str, ui_globs: list[str],
               non_ui_prefixes: list[str]) -> bool:
    """True si el path matchea un glob de UI Y no es un path no-UI."""
    # Excluir paths no-UI primero.
    for prefix in non_ui_prefixes:
        if prefix.endswith("/") and path.startswith(prefix):
            return False
        if not prefix.endswith("/") and (path == prefix or path.startswith(prefix + "/")):
            return False
    # Verificar si matchea un glob de UI.
    for glob in ui_globs:
        if fnmatch.fnmatch(path, glob):
            return True
    return False


def has_exemption_section(body: str, section: str) -> bool:
    """True si el cuerpo de la PR tiene la seccion de exencion con
    contenido no vacio."""
    lines = body.splitlines()
    in_section = False
    for line in lines:
        stripped = line.strip()
        lowered = stripped.lower()
        if lowered == section.lower() or lowered.startswith(section.lower() + " "):
            in_section = True
            # El resto de la linea despues del titulo.
            after = stripped[len(section):].strip()
            if after:
                return True
            continue
        if in_section:
            if stripped.startswith("#"):
                # Nueva seccion: la exencion estaba vacia.
                return False
            if stripped and not stripped.startswith("#"):
                return True
    return False


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="E2E-evidence gate (#425)",
    )
    parser.add_argument("--repo", default=".")
    parser.add_argument("--base", required=True,
                        help="Ref base de la comparacion")
    parser.add_argument("--head", default="HEAD",
                        help="Ref head de la comparacion (default: HEAD)")
    parser.add_argument("--ui-glob", action="append", default=None,
                        help="Glob de paths de UI; reemplaza los defaults, "
                             "puede repetirse")
    parser.add_argument("--non-ui-prefix", action="append", default=None,
                        help="Prefix no-UI; reemplaza la lista blanca "
                             "default, puede repetirse")
    parser.add_argument("--e2e-dir", default="tests/e2e_ci")
    parser.add_argument("--e2e-pattern", default="*.spec.{ts,js,py}")
    parser.add_argument("--pr-body-file", default=None,
                        help="Path al fichero con el cuerpo de la PR "
                             "(para verificar la seccion de exencion)")
    parser.add_argument("--exemption-section", default="## Exencion e2e")
    args = parser.parse_args(argv)

    repo = Path(args.repo).resolve()
    if not (repo / ".git").exists() and not (repo / ".git").is_file():
        print(f"FAIL: {repo} no es un repo git", file=sys.stderr)
        return 2

    ui_globs = args.ui_glob or DEFAULT_UI_GLOBS
    non_ui_prefixes = args.non_ui_prefix or DEFAULT_NON_UI

    try:
        changed = list_changed_paths(repo, args.base, args.head)
    except RuntimeError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2

    ui_changes = [p for p in changed
                  if is_ui_path(p, ui_globs, non_ui_prefixes)]

    if not ui_changes:
        print(
            f"OK: {len(changed)} paths cambiados, ninguno es UI; "
            f"e2e-evidence no aplica"
        )
        return 0

    # UI cambio: verificar si hay flujo e2e anadido.
    try:
        added_e2e = list_added_e2e_paths(
            repo, args.base, args.head, args.e2e_dir, args.e2e_pattern)
    except RuntimeError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2

    if added_e2e:
        print(
            f"OK: {len(ui_changes)} paths de UI cambiados, "
            f"{len(added_e2e)} flujo(s) e2e anadido(s) en {args.e2e_dir}"
        )
        return 0

    # Sin flujo e2e: verificar exencion.
    if args.pr_body_file:
        body_path = Path(args.pr_body_file)
        if not body_path.is_file():
            print(f"FAIL: cuerpo de la PR ausente: {args.pr_body_file}",
                  file=sys.stderr)
            return 2
        body = body_path.read_text(encoding="utf-8")
        if has_exemption_section(body, args.exemption_section):
            print(
                f"OK: {len(ui_changes)} paths de UI cambiados, sin e2e, "
                f"con seccion de exencion '{args.exemption_section}'"
            )
            return 0
    print(
        f"FAIL: {len(ui_changes)} paths de UI cambiados, sin flujo e2e "
        f"anadido en {args.e2e_dir} y sin seccion de exencion "
        f"'{args.exemption_section}' en el cuerpo de la PR",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
