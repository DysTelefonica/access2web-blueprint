#!/usr/bin/env bash
# Aprovisiona el directorio del runner nocturno (#1627, slice 6).
#
# Uso: provision-night-runner.sh --runner-dir DIR [--pins FICHERO] [--dry-run]
#
# 1. Comprueba cada herramienta del fichero de pines: debe existir y su versión debe
#    ser la fijada. Si falta o no cuadra, sale con 2 y lo nombra (no escribe nada).
# 2. Crea DIR y su marca DIR/.night-runner («profile=1»). Si la marca ya existe con
#    el mismo contenido no cambia nada; si existe con otro, sale con 2 sin tocarla.
# Idempotente: ejecutarlo dos veces deja el mismo estado. --dry-run solo imprime lo que
# haría y no escribe nada. Salida: 0 correcto, 2 algo no cuadra o uso no válido.
set -u

PROFILE_MARK="profile=1"
RUNNER_DIR=""
PINS="$(dirname "$0")/night-runner.pins"
DRY_RUN=0

while [ $# -gt 0 ]; do
  case "$1" in
    --runner-dir) [ $# -ge 2 ] || { echo "falta el valor de --runner-dir" >&2; exit 2; }
                  RUNNER_DIR="$2"; shift 2 ;;
    --pins)       [ $# -ge 2 ] || { echo "falta el valor de --pins" >&2; exit 2; }
                  PINS="$2"; shift 2 ;;
    --dry-run)    DRY_RUN=1; shift ;;
    *)            echo "argumento desconocido: $1" >&2; exit 2 ;;
  esac
done

if [ -z "$RUNNER_DIR" ]; then
  echo "falta --runner-dir" >&2
  exit 2
fi
if [ ! -f "$PINS" ]; then
  echo "no existe el fichero de pines: $PINS" >&2
  exit 2
fi

# 1. Herramientas y versiones.
problems=0
while IFS='=' read -r name expected; do
  case "$name" in ''|\#*) continue ;; esac
  if ! command -v "$name" >/dev/null 2>&1; then
    echo "falta la herramienta $name (versión fijada $expected en $PINS)" >&2
    problems=$((problems + 1))
    continue
  fi
  found="$("$name" --version 2>/dev/null | grep -oE '[0-9]+(\.[0-9]+)+' | head -1)"
  if [ "$found" != "$expected" ]; then
    echo "$name $found no es la versión fijada $expected en $PINS" >&2
    problems=$((problems + 1))
  fi
done < "$PINS"
if [ "$problems" -gt 0 ]; then
  echo "el runner nocturno no cuadra: $problems problema(s), no se ha escrito nada" >&2
  exit 2
fi

# 2. Directorio y marca.
MARK_FILE="$RUNNER_DIR/.night-runner"
if [ -f "$MARK_FILE" ]; then
  current="$(cat "$MARK_FILE")"
  if [ "$current" != "$PROFILE_MARK" ]; then
    echo "$MARK_FILE ya existe con «$current», no «$PROFILE_MARK»: no se sobrescribe" >&2
    exit 2
  fi
  if [ "$DRY_RUN" = 1 ]; then
    echo "ensayo: no hay nada que cambiar en $RUNNER_DIR (sin cambios)"
  else
    echo "runner nocturno ya provisionado en $RUNNER_DIR (sin cambios)"
  fi
  exit 0
fi

if [ "$DRY_RUN" = 1 ]; then
  echo "ensayo: crearía $RUNNER_DIR y $MARK_FILE; no se escribe nada"
  exit 0
fi

mkdir -p "$RUNNER_DIR" || { echo "no se puede crear $RUNNER_DIR" >&2; exit 2; }
printf '%s\n' "$PROFILE_MARK" > "$MARK_FILE" || { echo "no se puede escribir $MARK_FILE" >&2; exit 2; }
echo "runner nocturno provisionado en $RUNNER_DIR"
exit 0
