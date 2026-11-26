#!/usr/bin/env python3
"""night_gates.py: gate de ensayo nocturno en seco antes del merge (#1627, slice 8).

Gate autónomo: no está cableado en el merge gobernado (eso es el gate de parada de línea,
slice 2). Uso: night_gates.py dry-night --changed <ficheros> --head <sha> [--record <json>]
[--paths-file <json>] [--night-path <glob>]...

--changed: fichero de texto con una ruta por línea. Entrada documentada:
    git -c core.quotePath=false diff --no-renames --name-only <base>...<head>
Con core.quotePath=false git no cita los nombres no ASCII (una ruta citada sale con 2).
Con --no-renames un cambio de nombre trae la ruta vieja y la nueva. Una lista vacía
(o solo comentarios y líneas en blanco) sale con 2: un PR siempre cambia algo.
--record: JSON {"mode": "dry", "head_sha": "<sha>", "result": "success"}.
--paths-file: JSON con la lista de globs del pipeline nocturno; sustituye a NIGHT_PATHS.
--night-path: glob extra; repetible.

Salida: 0 si el PR puede fusionarse; 1 si toca el pipeline nocturno sin un ensayo en
seco en verde de su SHA; 2 si una entrada no se puede leer (fallo cerrado). Solo lee.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

# Globs del pipeline nocturno. `y*ml` cubre .yml y .yaml; `./` se normaliza antes.
NIGHT_PATHS = (
    ".github/workflows/nightly*.y*ml",
    ".github/workflows/night*.y*ml",
    "**/night-pipeline*.y*ml",
    "**/deploy-nightly*.y*ml",
    "**/ci-pattern/assets/prerelease/night*.py",
    "**/ci-pattern/assets/prerelease/night*.sh",
)

EXIT_OK, EXIT_NOT_VERIFIED, EXIT_UNREADABLE = 0, 1, 2
USAGE = ("uso: night_gates.py dry-night --changed <fichero> --head <sha> [--record <json>] "
         "[--paths-file <json>] [--night-path <glob>]...")


class PathsFileError(ValueError):
    def __init__(self, reason: str) -> None:
        super().__init__(f"el fichero de rutas no es válido: {reason}")


class EmptyPathsError(PathsFileError):
    def __init__(self) -> None:
        super().__init__("debe ser una lista JSON no vacía de globs")


class InvalidGlobError(PathsFileError):
    def __init__(self) -> None:
        super().__init__("cada entrada debe ser un glob no vacío")


class RecordShapeError(TypeError):
    def __init__(self) -> None:
        super().__init__("el registro debe ser un objeto JSON")


class UsageError(ValueError):
    def __init__(self, option: str) -> None:
        super().__init__(f"opción no reconocida o sin valor: {option}")


def _pin_output_encoding() -> None:
    """Fija stdout y stderr a UTF-8 antes de imprimir (#1648)."""
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def path_matches(path: str, pattern: str) -> bool:
    """Compara una ruta con un glob: `**` cruza directorios, `*` no."""
    regex = ""
    i = 0
    while i < len(pattern):
        if pattern.startswith("**/", i):
            regex += "(?:.*/)?"
            i += 3
        elif pattern.startswith("**", i):
            regex += ".*"
            i += 2
        elif pattern[i] == "*":
            regex += "[^/]*"
            i += 1
        elif pattern[i] == "?":
            regex += "[^/]"
            i += 1
        else:
            regex += re.escape(pattern[i])
            i += 1
    return re.fullmatch(regex, path) is not None


def _normalize(path: str) -> str:
    """Forma canónica: minúsculas, `/`, sin `/` ni `./` iniciales, sin `//` repetidos."""
    text = re.sub(r"/{2,}", "/", path.strip().replace("\\", "/").lower())
    while text.startswith(("./", "/")):
        text = text[2:] if text.startswith("./") else text[1:]
    return text


def touches_night_pipeline(changed: list[str], patterns=NIGHT_PATHS) -> list[str]:
    """Los ficheros cambiados que pertenecen al pipeline nocturno (sin distinguir mayúsculas)."""
    return [f for f in changed if any(path_matches(_normalize(f), p.lower()) for p in patterns)]


def check(changed: list[str], head: str, record: dict | None, patterns=NIGHT_PATHS):
    """Devuelve (código, mensaje). `record` es None si no hay registro."""
    touched = touches_night_pipeline(changed, patterns)
    if not touched:
        return EXIT_OK, "el PR no toca el pipeline nocturno: no hace falta ensayo"
    names = ", ".join(touched)
    if record is None:
        return EXIT_NOT_VERIFIED, (f"el PR toca el pipeline nocturno ({names}) y no hay un "
                                   f"ensayo nocturno registrado para {head}")
    if (record.get("mode") != "dry" or record.get("result") != "success"
            or record.get("head_sha") != head):
        return EXIT_NOT_VERIFIED, (f"el PR toca el pipeline nocturno ({names}) y el último "
                                   f"ensayo no es un ensayo en seco en verde de {head}")
    return EXIT_OK, f"ensayo nocturno en seco en verde para {head}"


class EmptyChangedError(ValueError):
    def __init__(self) -> None:
        super().__init__("la lista de cambios está vacía: un PR siempre cambia algo")


class QuotedPathError(ValueError):
    def __init__(self) -> None:
        super().__init__("git citó una ruta entre comillas (nombre no ASCII): ejecute "
                         "'git -c core.quotePath=false diff --no-renames --name-only <base>...<head>'")


def _read_changed(path: str) -> list[str]:
    """Rutas del PR. Un BOM inicial (PowerShell 5.1) se descarta; las líneas en blanco y los
    comentarios `#` no cuentan; una ruta citada por git sale con error (fallo cerrado)."""
    lines = [line.strip() for line in Path(path).read_text(encoding="utf-8-sig").splitlines()]
    if any(line.startswith('"') for line in lines):
        raise QuotedPathError()
    changed = [line for line in lines if line and not line.startswith("#")]
    if not changed:
        raise EmptyChangedError()
    return changed


def _read_record(path: str) -> dict:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise RecordShapeError()
    return data


def _read_paths(path: str) -> tuple[str, ...]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, list) or not data:
        raise EmptyPathsError()
    if not all(isinstance(glob, str) and glob for glob in data):
        raise InvalidGlobError()
    return tuple(data)


def _parse(args: list[str]) -> dict:
    options: dict[str, object] = {"--changed": None, "--head": None, "--record": None,
                                  "--paths-file": None, "--night-path": []}
    while args:
        flag = args.pop(0)
        if flag not in options or not args:
            raise UsageError(flag)
        if flag == "--night-path":
            options[flag].append(args.pop(0))  # type: ignore[union-attr]
        else:
            options[flag] = args.pop(0)
    return options


def main(argv: list[str]) -> int:
    _pin_output_encoding()
    if not argv or argv[0] != "dry-night":
        print(USAGE, file=sys.stderr)
        return EXIT_UNREADABLE
    try:
        options = _parse(argv[1:])
    except UsageError as error:
        print(f"{error}\n{USAGE}", file=sys.stderr)
        return EXIT_UNREADABLE
    if not options["--changed"] or not options["--head"]:
        print(f"faltan --changed o --head\n{USAGE}", file=sys.stderr)
        return EXIT_UNREADABLE
    try:
        changed = _read_changed(options["--changed"])
        record = _read_record(options["--record"]) if options["--record"] else None
        patterns = _read_paths(options["--paths-file"]) if options["--paths-file"] else NIGHT_PATHS
    except (OSError, ValueError, TypeError) as exc:
        print(f"no se puede leer la entrada: {exc}", file=sys.stderr)
        return EXIT_UNREADABLE
    patterns = (*patterns, *options["--night-path"])
    code, message = check(changed, options["--head"], record, patterns)
    stream = sys.stdout if code == EXIT_OK else sys.stderr
    print(message, file=stream)
    return code


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
