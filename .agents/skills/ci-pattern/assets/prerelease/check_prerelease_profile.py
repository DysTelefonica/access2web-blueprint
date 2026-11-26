#!/usr/bin/env python3
"""check_prerelease_profile.py: valida la declaración del perfil pre-release (#1627).

Uso: check_prerelease_profile.py <ruta/a/prerelease.json>

Forma aceptada. Acepta el fichero actual de APAP_WEB (`{"enabled": true}`); APAP_WEB
solo acepta `enabled`, así que allí no debe añadirse `schema_version` ni BOM:
  - enabled: obligatorio, booleano.
  - schema_version: opcional. Si está, entero >= 1 (ni float ni booleano). Si falta,
    vale 1.
  - Ningún campo fuera de esos dos.
  - Duplicar una clave es error. Un BOM UTF-8 se acepta.

Salida: 0 si la declaración es válida; 2 si no lo es, nombrando cada campo inválido
en stderr. Un fichero que no se puede leer, no es UTF-8 o no es JSON también sale con 2,
sin traceback (fallo cerrado).
"""

import json
import sys
from pathlib import Path

ALLOWED_FIELDS = {"schema_version", "enabled"}
USAGE_ARGC = 2  # programa + una ruta


def _pin_output_encoding():
    """Fija stdout y stderr a UTF-8 para que los mensajes con tildes salgan íntegros."""
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


class DuplicateKeyError(ValueError):
    """Clave repetida en un objeto JSON: error, no se pisa en silencio."""


def _reject_duplicate_keys(pairs):
    """object_pairs_hook: una clave repetida es un error."""
    result = {}
    for key, value in pairs:
        if key in result:
            raise DuplicateKeyError(key)
        result[key] = value
    return result


def validate(declaration):
    """Devuelve la lista de errores; vacía si la declaración es válida."""
    if not isinstance(declaration, dict):
        return ["declaración: debe ser un objeto JSON"]

    errors = []
    if "schema_version" in declaration:
        version = declaration["schema_version"]
        if isinstance(version, bool) or not isinstance(version, int) or version < 1:
            errors.append(f"schema_version: debe ser un entero >= 1, es {version!r}")

    if "enabled" not in declaration:
        errors.append("enabled: falta (booleano obligatorio)")
    elif not isinstance(declaration["enabled"], bool):
        errors.append(f"enabled: debe ser booleano, es {declaration['enabled']!r}")

    for field in sorted(set(declaration) - ALLOWED_FIELDS):
        errors.append(f"{field}: campo no admitido por el esquema")
    return errors


def load_declaration(path):
    """Devuelve (declaración, None) o (None, problema legible) sin lanzar excepciones."""
    try:
        raw = path.read_bytes()
    except OSError as exc:
        return None, f"no se puede leer {path}: {exc}"
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        return None, f"{path}: no es UTF-8 válido: {exc}"
    try:
        return json.loads(text, object_pairs_hook=_reject_duplicate_keys), None
    except DuplicateKeyError as exc:
        return None, f"{path}: clave duplicada «{exc.args[0]}»"
    except ValueError as exc:
        return None, f"{path}: JSON no válido: {exc}"


def main(argv):
    _pin_output_encoding()
    if len(argv) != USAGE_ARGC:
        print("uso: check_prerelease_profile.py <prerelease.json>", file=sys.stderr)
        return 2
    path = Path(argv[1])
    declaration, problem = load_declaration(path)
    if problem is not None:
        print(problem, file=sys.stderr)
        return 2

    errors = validate(declaration)
    for error in errors:
        print(f"{path}: {error}", file=sys.stderr)
    return 2 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
