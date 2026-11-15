#!/usr/bin/env python3
# check_deps_chain.py — el manifiesto y su lock cuentan lo mismo (HR-76, #1181)
"""La cadena de dependencias: el manifiesto que declara lo que se instala y el
lock que lo fija tienen que moverse juntos.

- Un PR que cambia un manifiesto de dependencias (el que declara
  ``[project].dependencies``, ``[project].optional-dependencies`` o
  ``dependency-groups``) y **no** cambia ninguno de sus locks es rojo: nombra el
  manifiesto, el lock que falta y el comando que lo regenera.
- Un repo sin lock, o un manifiesto que no declara dependencias, no da falso
  positivo: el gate avisa y sale 0.
- Un manifiesto ilegible es rojo (fail-closed, HR-3): un lock que no se puede
  medir no es un lock que se puede declarar conforme.

Por qué: Dependabot actualiza el manifiesto y no el lock, así que el patron de
la cadena se rompe en el PR que cambia el manifiesto sin regenerar el lock. El
CI instala el lock (``pip install --require-hashes -r app/requirements-dev.lock``)
y la imagen del PR se construye desde el manifiesto: con el lock atras, el CI
instala una version y la imagen lleva otra (``DysTelefonica/access2web-blueprint#1058``:
al subir SQLAlchemy a 2.1.6, que dejo de traer ``greenlet``, la imagen murio al
arrancar y ``e2e-smoke`` dio rojo).

Stdlib exclusivamente, sin red y sin ejecutar ``git``: recibe la lista de
ficheros cambiados del PR y lee el arbol.
"""

from __future__ import annotations

import sys
import tomllib
from pathlib import Path

#: Nombres de lock que dependen del manifiesto de su mismo directorio.
LOCK_NAMES = ("uv.lock", "requirements*.lock", "poetry.lock")

#: Comando canonico para re-resolver el lock de un directorio.
RELOCK_COMMAND = "uv lock --project <dir>"

#: Como se documenta el comando de exportacion de un lock.
EXPORT_COMMAND = ("uv export --project <dir> --frozen --no-emit-project --extra dev "
                   "--format requirements-txt --no-annotate --no-header")


class ManifestUnreadableError(RuntimeError):
    """El manifiesto no se puede leer o no es TOML valido: fallo cerrado (HR-3).

    El mensaje lo compone la clase (TRY003): quien lanza dice QUE no se pudo
    leer y por que, no el texto entero.
    """

    def __init__(self, path: str, cause: object) -> None:
        super().__init__(f"{path}: {cause}")


class ChangedFilesUnreadableError(RuntimeError):
    """La lista de ficheros cambiados del PR no se puede leer: fallo cerrado (HR-3)."""

    def __init__(self, path: str, cause: object) -> None:
        super().__init__(f"no se puede leer la lista de ficheros cambiados ({path}): {cause}")


def declares_dependencies(manifest: Path) -> bool:
    """True cuando el manifiesto declara dependencias que instalar.

    Un ``pyproject.toml`` que solo lleva configuracion de herramientas (el caso
    del manifiesto raiz de un consumer) no es un manifiesto de dependencias: no
    tiene lock y por tanto no genera hallazgo.
    """
    try:
        data = tomllib.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise ManifestUnreadableError(str(manifest), exc) from exc
    project = data.get("project")
    project = project if isinstance(project, dict) else {}
    return bool(project.get("dependencies") or project.get("optional-dependencies")
                or data.get("dependency-groups"))


def locks_of(manifest: Path) -> list[Path]:
    """Los locks que dependen del manifiesto: los de su mismo directorio."""
    directory = manifest.parent
    found: list[Path] = []
    for name in LOCK_NAMES:
        if "*" in name:
            for entry in sorted(directory.glob(name)):
                if entry.is_file():
                    found.append(entry)
        else:
            candidate = directory / name
            if candidate.is_file():
                found.append(candidate)
    return found


def read_changed_files(path: Path) -> list[str]:
    """Los ficheros cambiados del PR, uno por linea (los da el workflow)."""
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise ChangedFilesUnreadableError(str(path), exc) from exc
    return [line.strip() for line in text.splitlines() if line.strip()]


def manifests_of(repo: Path) -> list[Path]:
    """Los manifiestos de dependencias del repo, ordenados por ruta."""
    found: list[Path] = []
    for path in sorted(repo.rglob("pyproject.toml")):
        if any(part in (".venv", "node_modules", ".git") for part in path.parts):
            continue
        found.append(path)
    return found


def _changed_manifests(repo: Path, changed_files: list[str]) -> list[Path]:
    """Los manifiestos de dependencias que el PR cambia."""
    changed = {Path(item).as_posix() for item in changed_files}
    out: list[Path] = []
    for manifest in manifests_of(repo):
        if manifest.relative_to(repo).as_posix() not in changed:
            continue
        if declares_dependencies(manifest):
            out.append(manifest)
    return out


def check(repo: Path, changed_files: list[str]) -> list[str]:
    """Hallazgos de la cadena manifiesto ↔ lock. Vacío cuando todo cuadra.

    Un manifiesto sin lock a su lado no es un hallazgo: el gate no puede exigir
    una coherencia que el consumer no ha declarado (ver :func:`notes`).
    """
    out: list[str] = []
    for manifest in _changed_manifests(repo, changed_files):
        locks = locks_of(manifest)
        if not locks:
            continue
        relative = manifest.relative_to(repo).as_posix()
        if not any(lock.relative_to(repo).as_posix() in changed_files for lock in locks):
            lock_names = ", ".join(lock.relative_to(repo).as_posix() for lock in locks)
            directory = manifest.parent.relative_to(repo).as_posix()
            out.append(
                f"HR-76: el manifiesto '{relative}' cambio y su lock no ({lock_names}): "
                f"el CI instala el lock y la imagen se construye desde el manifiesto, "
                f"asi que el bump no se aplicaria. Regenere el lock: "
                f"{RELOCK_COMMAND.replace('<dir>', directory)} y "
                f"{EXPORT_COMMAND.replace('<dir>', directory)}")
    return out


def notes(repo: Path, changed_files: list[str]) -> list[str]:
    """Avisos: lo que el gate no puede medir (un manifiesto sin lock)."""
    out: list[str] = []
    for manifest in _changed_manifests(repo, changed_files):
        if not locks_of(manifest):
            relative = manifest.relative_to(repo).as_posix()
            out.append(f"AVISO: el manifiesto '{relative}' declara dependencias y no tiene "
                       f"lock a su lado ({', '.join(LOCK_NAMES)}): la cadena no se mide")
    return out


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python saca la codificacion de la plataforma, asi que el mismo gate
    emite bytes distintos en un puesto Windows (cp1252) y en un runner Linux
    (utf-8), y un caracter que no se puede codificar revienta la escritura.
    """
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    import argparse

    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", default=".", help="raiz del repositorio")
    parser.add_argument("--changed-files", required=True,
                        help="fichero con los caminos cambiados del PR, uno por linea")
    args = parser.parse_args(argv)
    try:
        changed = read_changed_files(Path(args.changed_files))
        repo = Path(args.repo)
        findings = check(repo, changed)
        for item in notes(repo, changed):
            print(item, file=sys.stderr)
    except (ManifestUnreadableError, ChangedFilesUnreadableError) as exc:
        print(f"HR-76: {exc}", file=sys.stderr)
        return 1
    for item in findings:
        print(item, file=sys.stderr)
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
