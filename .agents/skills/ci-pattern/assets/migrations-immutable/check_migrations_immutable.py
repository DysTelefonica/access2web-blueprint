#!/usr/bin/env python3
# ci-pattern asset — migrations-immutable gate (DysTelefonica/team-skills#431)
"""Gate de inmutabilidad de migraciones versionadas.

Caso de uso: consumers web con migraciones SQL versionadas (nomenclatura
``NNN_descripcion.sql``). El runner registra las migraciones aplicadas
por nombre de fichero; renombrar, mover, borrar o editar una migracion
ya aplicada en produccion provoca que el siguiente deploy re-ejecute su
DDL o la pierda.

El gate compara el arbol de migraciones de HEAD contra la rama base:

- **Rename** (migracion presente en base con path distinto en HEAD) -> rojo.
- **Delete** (migracion presente en base, ausente en HEAD) -> rojo.
- **Edit** (migracion con mismo path, contenido distinto) -> rojo.
- **New** (migracion nueva en HEAD) -> requiere prefijo numerico mayor
  al maximo de la base.
- **Duplicate prefix** (prefijo ya duplicado en la base) -> tolerado;
  uno nuevo con prefijo duplicado -> rojo.
- **Reserved prefix** (prefijo ya en una PR abierta contra main) -> rojo.
  Caso real APAP_WEB 2026-10-13 (#431, comentario del mantenedor): tres
  PR abiertas a la vez (#1465, #1565, #1573) anadian cada una su propia
  ``011_*.sql``; cada PR pasaba el CI por separado pero al mergear la
  segunda habria colision de numeracion. Cubre el hueco que el gate de
  rama base no ve.

Parametros del consumer (CLI):

- ``--migrations-dir <path>``: directorio de migraciones relativo a la raiz
  del consumer (default: ``migrations``).
- ``--pattern <glob>``: patron de ficheros de migracion (default:
  ``*.sql``). Sin recursion; el gate no asume jerarquia.
- ``--base <ref>``: ref base de la comparacion (default: ``origin/main``).
  Cualquier SHA, branch o tag que ``git rev-parse`` resuelva.
- ``--no-check-reserved-prefixes``: opt-out del check de PRs abiertas
  (caso del consumer que aun no migra a la API). Default ON.

Exit codes:

- 0: limpio.
- 1: hallazgos (rename, delete, edit, new con prefijo <= maximo, duplicate
  prefix nuevo, reserved prefix).
- 2: entrada invalida (dir ausente, git falla, patron no resuelve
  ficheros, base no resuelta).
- 3: el check de reservados fallo por error de API (fail-closed).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

# Prefijo numerico al inicio del nombre: N, NN, NNN, etc. (sin limite de
# digitos: 9999 sigue siendo valido si el consumer usa 4 digitos).
PREFIX_RE = re.compile(r"^(\d+)[_-]")

# API REST base. Override con GITHUB_API_BASE para tests que mockean
# la respuesta (un http.server local, por ejemplo).
GITHUB_API_BASE = os.environ.get("GITHUB_API_BASE", "https://api.github.com")


def run_git(repo: Path, *args: str) -> str:
    """Ejecuta git en ``repo`` y devuelve stdout. Non-zero rc -> RuntimeError."""
    proc = subprocess.run(
        ["git", "-C", str(repo), *args],
        capture_output=True, text=True, check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"git {' '.join(args)} fallo (rc={proc.returncode}): {proc.stderr.strip()}"
        )
    return proc.stdout


def list_migrations(repo: Path, migrations_dir: str, pattern: str) -> dict[str, str]:
    """Lista las migraciones de ``migrations_dir`` en ``ref``.

    Devuelve ``{relative_path: blob_sha}``. Paths relativos a la raiz
    del repo (no al cwd) para que el compare con base sea directo.
    """
    out = run_git(
        repo, "ls-tree", "-r", "HEAD", "--", migrations_dir,
    )
    result = {}
    for line in out.splitlines():
        # formato: <mode> <type> <sha>\t<path>. El path puede tener
        # espacios; el tab es el separador fiable entre sha y path.
        if "\t" not in line:
            continue
        meta_and_sha, rel = line.split("\t", 1)
        sha = meta_and_sha.split()[-1]
        if not rel:
            continue
        # Filtrar por patron (fnmatch basico: solo soporta * y ?)
        import fnmatch
        if not fnmatch.fnmatch(Path(rel).name, pattern):
            continue
        result[rel] = sha
    return result


def parse_prefix(filename: str) -> int | None:
    """Extrae el prefijo numerico del nombre de fichero. None si no
    tiene prefijo (migracion sin nomenclatura versionada)."""
    m = PREFIX_RE.match(Path(filename).name)
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------------------------
# Reserved prefixes (open PRs against main)
# ---------------------------------------------------------------------------------------------


def _http_get_json(url: str, token: str) -> object:
    """GET ``url`` con headers de GitHub y devuelve el JSON parseado.

    Stdlib only. Timeout 15s. Los HTTPError se reempaquetan en un
    RuntimeError con el codigo y la causa para que el caller decida
    si falla cerrado.
    """
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/vnd.github+json", "User-Agent": "ci-pattern/migrations-immutable"},
    )
    if token:
        request.add_header("Authorization", f"Bearer {token}")
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"GitHub API returned HTTP {exc.code} for {url}: {exc.reason}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"GitHub API request failed for {url}: {exc.reason}") from exc


def parse_repo_slug(repo: Path) -> str | None:
    """Lee ``origin`` del repo y devuelve ``owner/name``.

    Acepta URLs SSH (``git@github.com:owner/name.git``) y HTTPS
    (``https://github.com/owner/name(.git)``). Devuelve None si el
    remote no es GitHub o no se puede parsear — el caller falla
    cerrado con un mensaje claro.
    """
    try:
        out = run_git(repo, "config", "--get", "remote.origin.url")
    except RuntimeError:
        return None
    url = out.strip()
    # SSH: git@github.com:owner/name.git. El prefijo exacto ya descarta
    # hosts maliciosos: nada puede inyectar ``github.com`` antes de los
    # dos puntos sin dejar de parecerse a una URL SSH de GitHub.
    if url.startswith("git@github.com:"):
        path = url.split(":", 1)[1]
    else:
        # HTTPS y similares: parseamos con urllib.parse y exigimos
        # ``hostname == "github.com"`` exacto. Un match por substring
        # (``"github.com/" in url``) acepta URLs como
        # ``https://evil.example/github.com/owner/repo.git`` y
        # ``https://github.com.evil.com/owner/repo.git``, dejando al
        # caller enviar la peticion REST con un slug arbitrario.
        # CodeQL ``py/incomplete-url-substring-sanitization`` marco la
        # linea 159 con esta clase de bug; PR #910 del blueprint
        # origino la remediacion que propagamos aca (Refs #528).
        parsed = urllib.parse.urlparse(url)
        if parsed.hostname != "github.com":
            return None
        # ``parsed.path`` arranca con ``/``; lo quitamos para que el
        # split final (``owner/repo`` -> [owner, repo]) coincida con
        # el de la rama SSH.
        path = parsed.path.lstrip("/")
    path = path.rstrip("/")
    if path.endswith(".git"):
        path = path[:-4]
    parts = path.split("/")
    if len(parts) != 2 or not all(parts):
        return None
    return f"{parts[0]}/{parts[1]}"


def _pr_files_status_is_added(file_entry: dict) -> bool:
    """True cuando el fichero fue anadido por la PR (``status == 'added'``).

    Tambien cubre el caso ``status == 'renamed'`` y ``previous_filename``
    presente: si la PR lo renombro desde fuera de migrations_dir, lo
    tomamos como anadido al directorio de migraciones.
    """
    return file_entry.get("status") in {"added", "renamed"}


def fetch_open_pr_migration_prefixes(
    slug: str,
    token: str,
    migrations_dir: str,
    pattern: str,
) -> tuple[set[int], dict[int, int]]:
    """Lista los prefijos de migraciones anadidas en PRs abiertas contra main.

    Devuelve ``(reserved_prefixes, reserved_sources)`` donde
    ``reserved_sources[prefix] = pr_number`` que reservo el prefijo.
    Solo se cuentan los ficheros en ``migrations_dir`` que matcheen
    ``pattern`` y tengan un prefijo numerico extraible.

    Lanza RuntimeError si la API falla; el caller decide si falla
    cerrado (default).
    """
    import fnmatch
    reserved: set[int] = set()
    sources: dict[int, int] = {}

    pulls_url = f"{GITHUB_API_BASE}/repos/{slug}/pulls?state=open&base=main&per_page=100"
    pulls = _http_get_json(pulls_url, token)
    if not isinstance(pulls, list):
        raise RuntimeError(f"GitHub API returned non-list for {pulls_url}: {type(pulls).__name__}")

    for pr in pulls:
        pr_number = pr.get("number")
        if not isinstance(pr_number, int):
            continue
        files_url = f"{GITHUB_API_BASE}/repos/{slug}/pulls/{pr_number}/files?per_page=100"
        files = _http_get_json(files_url, token)
        if not isinstance(files, list):
            continue
        for file_entry in files:
            if not isinstance(file_entry, dict):
                continue
            if not _pr_files_status_is_added(file_entry):
                continue
            path = file_entry.get("filename", "")
            # Solo ficheros bajo migrations_dir.
            if not path.startswith(migrations_dir + "/") and path != migrations_dir:
                continue
            if not fnmatch.fnmatch(Path(path).name, pattern):
                continue
            prefix = parse_prefix(path)
            if prefix is None:
                continue
            reserved.add(prefix)
            sources.setdefault(prefix, pr_number)

    return reserved, sources


def check_reserved_prefix_conflict(
    new_prefixes: set[int],
    reserved_prefixes: set[int],
    reserved_sources: dict[int, int],
    migrations_dir: str,
) -> list[str]:
    """Devuelve findings ``reserved: prefijo N ya en PR #X`` para cada
    colision entre prefijos nuevos y reservados.
    """
    findings: list[str] = []
    for prefix in sorted(new_prefixes & reserved_prefixes):
        pr_number = reserved_sources.get(prefix, "?")
        findings.append(
            f"reserved: prefijo {prefix:03d} ya esta en una PR abierta "
            f"contra main (#{pr_number}); dos PRs con el mismo prefijo "
            f"colisionan al mergear la segunda. Usa el siguiente "
            f"prefijo libre o coordina con quien tiene la PR #{pr_number}."
        )
    return findings


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str]) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(
        description="Migrations-immutable gate (#431)",
    )
    parser.add_argument(
        "--repo", default=".",
        help="Ruta al repo del consumer (default: cwd)",
    )
    parser.add_argument(
        "--migrations-dir", default="migrations",
        help="Directorio de migraciones relativo al repo (default: migrations)",
    )
    parser.add_argument(
        "--pattern", default="*.sql",
        help="Patron glob de ficheros de migracion (default: *.sql)",
    )
    parser.add_argument(
        "--base", default="origin/main",
        help="Ref base de la comparacion (default: origin/main)",
    )
    parser.add_argument(
        "--no-check-reserved-prefixes",
        action="store_true",
        help=(
            "Desactiva el check de PRs abiertas contra main (caso del "
            "consumer que aun no migra a la API). Default: activado."
        ),
    )
    args = parser.parse_args(argv)

    repo = Path(args.repo).resolve()

    # Validar que el repo es un repo git.
    if not (repo / ".git").exists() and not (repo / ".git").is_file():
        print(f"FAIL: {repo} no es un repo git", file=sys.stderr)
        return 2

    try:
        return _run(args, repo)
    except RuntimeError as exc:
        # git fallo (base no resuelta, ls-tree de dir inexistente, etc).
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2


def _run(args, repo: Path) -> int:

    # Validar que el dir de migraciones existe en HEAD.
    head_files = list_migrations(repo, args.migrations_dir, args.pattern)

    # Listar migraciones en base. git ls-tree no falla si el dir no existe
    # (devuelve vacio); detectamos eso como dir ausente.
    base_out = run_git(
        repo, "ls-tree", "-r", args.base, "--", args.migrations_dir,
    )
    base_files: dict[str, str] = {}
    import fnmatch
    for line in base_out.splitlines():
        if "\t" not in line:
            continue
        meta_and_sha, rel = line.split("\t", 1)
        sha = meta_and_sha.split()[-1]
        if not fnmatch.fnmatch(Path(rel).name, args.pattern):
            continue
        base_files[rel] = sha

    if not head_files and not base_files:
        # El dir no existe en ninguna de las dos: entorno sin
        # migraciones, no es nuestro caso.
        if not (repo / args.migrations_dir).exists():
            pass  # OK, el consumer no tiene migraciones
        else:
            print(
                f"FAIL: {args.migrations_dir} existe pero no contiene ficheros "
                f"que matcheen {args.pattern!r}",
                file=sys.stderr,
            )
            return 2

    findings: list[str] = []

    # Detectar renames y deletes.
    head_paths = set(head_files.keys())
    base_paths = set(base_files.keys())

    # Rename: misma sha, distinto path.
    base_by_sha: dict[str, str] = {sha: path for path, sha in base_files.items()}
    for head_path, head_sha in head_files.items():
        if head_sha in base_by_sha and base_by_sha[head_sha] != head_path:
            findings.append(
                f"rename: {base_by_sha[head_sha]} -> {head_path} "
                f"(migracion existente con sha {head_sha[:7]} movida de sitio)"
            )

    # Delete: en base, no en head.
    for base_path in base_paths:
        if base_path not in head_paths:
            # Excluir renames ya reportados (la sha esta en head con otro path).
            base_sha = base_files[base_path]
            if base_sha not in head_files.values():
                findings.append(
                    f"delete: {base_path} (migracion existente borrada)"
                )

    # Edit: mismo path, sha distinta.
    for path in head_paths & base_paths:
        if head_files[path] != base_files[path]:
            findings.append(
                f"edit: {path} (migracion existente modificada)"
            )

    # New: en head, no en base. Validar prefijo.
    new_paths = head_paths - base_paths
    base_prefixes = [parse_prefix(p) for p in base_paths]
    base_prefixes = [p for p in base_prefixes if p is not None]
    max_base_prefix = max(base_prefixes) if base_prefixes else 0

    # Detectar duplicados historicos en la base (tolerados).
    base_prefix_counts: dict[int, int] = defaultdict(int)
    for p in base_prefixes:
        base_prefix_counts[p] += 1
    historical_duplicates = {k for k, v in base_prefix_counts.items() if v > 1}

    for new_path in sorted(new_paths):
        prefix = parse_prefix(new_path)
        if prefix is None:
            findings.append(
                f"new: {new_path} (sin prefijo numerico; el gate exige "
                f"patron NNN_descripcion)"
            )
            continue
        if prefix in historical_duplicates:
            findings.append(
                f"new: {new_path} (prefijo {prefix} ya duplicado en la base; "
                f"el gate tolera duplicados historicos pero rechaza nuevos)"
            )
            continue
        if prefix <= max_base_prefix:
            findings.append(
                f"new: {new_path} (prefijo {prefix} <= maximo de la base "
                f"{max_base_prefix}; las nuevas deben ir despues)"
            )
            continue

    # Reserved prefixes: prefijos que ya estan en una PR abierta contra
    # main. Caso real APAP_WEB 2026-10-13: tres PR anadian su propia
    # ``011_*.sql`` (#1465, #1565, #1573) y el gate de rama base no
    # las detectaba. Este check cierra el hueco. Fail-closed: si el
    # check esta activo y no podemos medir, salimos en rojo.
    if not args.no_check_reserved_prefixes and new_paths:
        new_prefixes = {
            parse_prefix(p) for p in new_paths
            if parse_prefix(p) is not None
        }
        if new_prefixes:
            token = os.environ.get("GITHUB_TOKEN", "").strip()
            if not token:
                print(
                    "FAIL: --no-check-reserved-prefixes no esta activado pero "
                    "GITHUB_TOKEN no esta en el entorno; no se puede "
                    "comprobar si los prefijos nuevos estan reservados por "
                    "otras PRs abiertas. Configura GITHUB_TOKEN o pasa "
                    "--no-check-reserved-prefixes para opt-out explicito.",
                    file=sys.stderr,
                )
                return 3
            slug = parse_repo_slug(repo)
            if not slug:
                print(
                    f"FAIL: no se pudo extraer owner/name del remote "
                    f"origin del repo {repo}; el check de reservados "
                    f"necesita el slug del repo para consultar la API.",
                    file=sys.stderr,
                )
                return 3
            try:
                reserved_prefixes, reserved_sources = fetch_open_pr_migration_prefixes(
                    slug=slug,
                    token=token,
                    migrations_dir=args.migrations_dir,
                    pattern=args.pattern,
                )
            except RuntimeError as exc:
                print(f"FAIL: {exc}", file=sys.stderr)
                return 3
            findings.extend(
                check_reserved_prefix_conflict(
                    new_prefixes=new_prefixes,
                    reserved_prefixes=reserved_prefixes,
                    reserved_sources=reserved_sources,
                    migrations_dir=args.migrations_dir,
                )
            )

    if findings:
        print(f"FAIL: {len(findings)} hallazgo(s) en migraciones:")
        for f in findings:
            print(f"  {f}")
        return 1

    print(
        f"OK: {len(head_files)} migracion(es) en HEAD, "
        f"{len(base_files)} en {args.base}, sin hallazgos"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
