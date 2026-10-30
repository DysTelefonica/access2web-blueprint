#!/usr/bin/env python3
"""Comprueba el árbol vendorizado del patrón y su sello contra el commit base (#1261).

En un job de `pull_request` no hay secreto para leer el catálogo, así que el
gate no puede medir contra el canon vivo. Lo que sí puede: que el PR no cambie
el código que mide. Este script compara el PR con su base:

1. El sello `catalog @ <sha>` de la primera línea del manifiesto debe existir
   en la base, con un sha hexadecimal de 7 a 40 caracteres (HR-3: sin sello no
   hay commit que medir, así que falla cerrado).
2. La primera línea del manifiesto del PR debe ser idéntica a la de la base:
   el PR no puede reapuntar el gate a otro commit del catálogo.
3. El árbol vendorizado del PR debe ser idéntico al de la base, incluidos los
   ficheros sin trackear. Un PR que altera el gate no pasa.

Lo que no puede comprobar: que el árbol vendorizado coincida con el catálogo en
el sello. Eso exige leer el catálogo, y eso exige el secreto, que este job no
recibe. Esa comprobación queda en un camino de confianza, fuera de los PR.

Se ejecuta desde la copia de la base del propio script (`git show <base>:...`),
así que un PR no puede editar el comprobador para pasar.

Uso:
    vendored_seal_check.py --base <sha> --vendored <ruta> --manifest <ruta> [--repo <dir>]

Salida: 0 si el PR coincide con su base; 1 en cualquier otro caso.
Solo stdlib, sin red.
"""
import argparse
import re
import subprocess
import sys
from pathlib import Path

SEAL_RE = re.compile(r"catalog @ ([0-9a-f]{7,40})\)")


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def git(repo: Path, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(repo), *args],  # noqa: S603, S607 — git con argumentos fijos, sin shell.
                          capture_output=True, text=True)


def first_line_at(repo: Path, rev: str, manifest: str) -> str | None:
    """Primera línea del manifiesto en `rev`, o None si `rev` o el fichero no existen."""
    proc = git(repo, "show", f"{rev}:{manifest}")
    if proc.returncode != 0:
        return None
    return proc.stdout.splitlines()[0] if proc.stdout else ""


def check(repo: Path, base: str, vendored: str, manifest: str) -> list[str]:
    """Problemas encontrados; lista vacía si el PR coincide con su base."""
    base_seal = first_line_at(repo, base, manifest)
    if base_seal is None:
        return [f"no se puede leer {manifest} en la base {base}: sin commit base no hay sello que medir (HR-3)"]
    if not SEAL_RE.search(base_seal):
        return [f"la base de {manifest} no trae un sello 'catalog @ <sha>' hexadecimal: no hay commit del patrón que medir (HR-3)"]

    head_manifest = (repo / manifest)
    head_seal = head_manifest.read_text(encoding="utf-8").splitlines()[0] if head_manifest.is_file() else ""
    problems = []
    if head_seal != base_seal:
        problems.append(f"{manifest} cambia el sello respecto a la base: el PR no puede reapuntar el gate a otro commit del catálogo")

    changed = git(repo, "diff", "--name-only", base, "--", vendored).stdout.split()
    untracked = git(repo, "ls-files", "--others", "--exclude-standard", "--", vendored).stdout.split()
    for path in changed + untracked:
        problems.append(f"el PR altera el árbol vendorizado {vendored}: {path}")
    return problems


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--base", required=True, help="sha del commit base del PR (base.sha)")
    parser.add_argument("--vendored", required=True, help="ruta del árbol vendorizado del patrón")
    parser.add_argument("--manifest", required=True, help="ruta del manifiesto con el sello")
    parser.add_argument("--repo", default=".", help="directorio del checkout del PR")
    args = parser.parse_args(argv)

    problems = check(Path(args.repo), args.base, args.vendored, args.manifest)
    if problems:
        for problem in problems:
            print(f"FAIL: {problem}")
        return 1
    print(f"OK: {args.vendored} y el sello de {args.manifest} coinciden con la base {args.base}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
