#!/bin/sh
# tooling-preflight.sh — preflight fail-loud del contrato de tooling del job
# (skill oracle-vps-github-runners, reglas 16/17; hueco G4 del inventario de
# activos, `references/asset-inventory.md`; issue DysTelefonica/team-skills#165).
#
# Regla: todo job migrado declara su contrato de tooling completo (cada
# binario que invocan sus `run:` y contenedores) como DATO del job, y este
# preflight fail-loud al inicio del job verifica que cada binario existe en
# el runner que lo ejecuta. Verificarlo solo en el runner histórico no prueba
# nada: la ausencia se demuestra en el primero que no lo tiene. Un preflight
# sin lista declarada falla; NUNCA pasa en silencio.
#
# Uso:
#   tooling-preflight.sh --binaries FICHERO
#
# FICHERO es el dato versionado del job (en el repo del consumer, junto a la
# definición del servicio — nunca en el host): una entrada por línea, `#`
# inicia comentario, las líneas en blanco se ignoran.
#
# Exit codes:
#   0  todos los binarios declarados están presentes en este runner;
#   1  al menos uno ausente — se nombran TODOS por stderr;
#   2  entrada inválida: falta --binaries, la lista no existe, no es legible
#      o está vacía (un contrato sin declarar es un fallo, no un pase).

set -u

PROGNAME=${0##*/}

usage_fail() {
    echo "$PROGNAME: $*" >&2
    echo "Uso: $PROGNAME --binaries FICHERO (una entrada por línea; '#' comentario; línea en blanco ignorada)" >&2
    exit 2
}

binaries_file=""
while [ $# -gt 0 ]; do
    case $1 in
        --binaries)
            [ $# -ge 2 ] || usage_fail "la opción --binaries requiere un fichero"
            binaries_file=$2
            shift 2
            ;;
        *)
            usage_fail "argumento no reconocido: $1"
            ;;
    esac
done

[ -n "$binaries_file" ] || usage_fail "falta --binaries: sin lista declarada el preflight no puede pasar"

if [ ! -f "$binaries_file" ] || [ ! -r "$binaries_file" ]; then
    echo "$PROGNAME: la lista de binarios declarada no existe o no es legible: $binaries_file" >&2
    exit 2
fi

declared=0
entries=""
while IFS= read -r line || [ -n "$line" ]; do
    # Recorte de comentario y espacios en blanco extremos.
    entry=$(printf '%s' "$line" | sed -e 's/#.*//' -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
    [ -n "$entry" ] || continue
    if [ -z "$entries" ]; then
        entries=$entry
    else
        entries="$entries
$entry"
    fi
    declared=$((declared + 1))
done < "$binaries_file"

if [ "$declared" -eq 0 ]; then
    echo "$PROGNAME: la lista declarada está vacía: el contrato de tooling no está declarado y el preflight falla, no pasa en silencio" >&2
    exit 2
fi

missing=""
missing_count=0
while IFS= read -r entry; do
    [ -n "$entry" ] || continue
    if ! command -v "$entry" >/dev/null 2>&1; then
        if [ -z "$missing" ]; then
            missing=$entry
        else
            missing="$missing
$entry"
        fi
        missing_count=$((missing_count + 1))
    fi
done <<EOF
$entries
EOF

if [ "$missing_count" -gt 0 ]; then
    echo "$PROGNAME: contrato de tooling incumplido en el runner que ejecuta este job ($missing_count de $declared ausentes):" >&2
    printf '%s\n' "$missing" >&2
    echo "$PROGNAME: el fix aterriza en la definición del servicio, nunca en una instalación manual del host (regla 17)" >&2
    exit 1
fi

echo "$PROGNAME: contrato de tooling verificado en este runner ($declared binarios presentes)"
exit 0
