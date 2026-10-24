#!/bin/sh
# tooling-preflight.sh — preflight fail-loud del contrato de tooling del job
# (skill oracle-vps-github-runners, reglas 16/17; hueco G4 del inventario de
# activos, `references/asset-inventory.md`; issue DysTelefonica/team-skills#165;
# identidad de runner, TS#1113 clase M).
#
# Regla: todo job migrado declara su contrato de tooling completo (cada
# binario que invocan sus `run:` y contenedores) como DATO del job, y este
# preflight fail-loud al inicio del job verifica que cada binario existe en
# el runner que lo ejecuta. Verificarlo solo en el runner histórico no prueba
# nada: la ausencia se demuestra en el primero que no lo tiene. Un preflight
# sin lista declarada falla; NUNCA pasa en silencio.
#
# Identidad del runner: `--runner <nombre>` declara en qué runner se ejecuta
# el job y es lo que hace accionable el fallo: sin ella, «falta el binario»
# no dice dónde mirar (el fix aterriza en la definición del servicio del
# runner concreto, nunca en una instalación manual de un host cualquiera).
# La opción es opt-in: la invocación heredada `--binaries FICHERO` a secas
# conserva su comportamiento observable (mismos códigos y mismo texto de
# ausencia); lo que añade `--runner` es el mensaje exacto por CLI ausente.
#
# Uso:
#   tooling-preflight.sh --binaries FICHERO [--runner NOMBRE]
#
# FICHERO es el dato versionado del job (en el repo del consumer, junto a la
# definición del servicio — nunca en el host): una entrada por línea, `#`
# inicia comentario, las líneas en blanco se ignoran.
# NOMBRE es la identidad del runner que ejecuta el job: en el workflow del
# consumer, `${{ runner.name }}`.
#
# Exit codes:
#   0  todos los binarios declarados están presentes en este runner;
#   1  al menos uno ausente — por stderr, una línea por ausente:
#      `falta <binario> en el runner <nombre>` si la invocación declaró
#      `--runner`; en la invocación heredada, el resumen genérico con la
#      lista de ausentes;
#   2  entrada inválida: falta --binaries, la lista no existe, no es legible
#      o está vacía (un contrato sin declarar es un fallo, no un pase),
#      `--runner` sin nombre, o hay una opción no reconocida. La AUSENCIA de
#      `--runner` NO es una entrada inválida: es la invocación heredada.

set -u

PROGNAME=${0##*/}

usage_fail() {
    echo "$PROGNAME: $*" >&2
    echo "Uso: $PROGNAME --binaries FICHERO [--runner NOMBRE] (una entrada por línea; '#' comentario; línea en blanco ignorada)" >&2
    exit 2
}

binaries_file=""
runner_name=""
runner_given=0
while [ $# -gt 0 ]; do
    case $1 in
        --binaries)
            [ $# -ge 2 ] || usage_fail "la opción --binaries requiere un fichero"
            binaries_file=$2
            shift 2
            ;;
        --runner)
            [ $# -ge 2 ] || usage_fail "la opción --runner requiere un nombre de runner"
            runner_given=1
            runner_name=$2
            shift 2
            ;;
        *)
            usage_fail "argumento no reconocido: $1"
            ;;
    esac
done

[ -n "$binaries_file" ] || usage_fail "falta --binaries: sin lista declarada el preflight no puede pasar"

# La identidad del runner se normaliza (espacios extremos fuera) para que el
# mensaje de ausencia sea exactamente `falta <cli> en el runner <nombre>`.
# Declarada pero vacía es una invocación inválida; no declarada es la
# invocación heredada, que sigue siendo válida.
if [ "$runner_given" -eq 1 ]; then
    runner_name=$(printf '%s' "$runner_name" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
    [ -n "$runner_name" ] || usage_fail "--runner sin nombre: la identidad del runner no puede estar vacía"
fi

if [ ! -f "$binaries_file" ] || [ ! -r "$binaries_file" ]; then
    echo "$PROGNAME: la lista de binarios declarada no existe o no es legible: $binaries_file" >&2
    exit 2
fi

# Una sola pasada sobre la lista: se cuentan los declarados y se guardan los
# ausentes (en el orden de la declaración). La lista vacía no es un pase:
# sale 2.
declared=0
missing_count=0
missing=""
while IFS= read -r line || [ -n "$line" ]; do
    # Recorte de comentario y espacios en blanco extremos.
    entry=$(printf '%s' "$line" | sed -e 's/#.*//' -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
    [ -n "$entry" ] || continue
    declared=$((declared + 1))
    if ! command -v "$entry" >/dev/null 2>&1; then
        if [ -z "$missing" ]; then
            missing=$entry
        else
            missing="$missing
$entry"
        fi
        missing_count=$((missing_count + 1))
    fi
done < "$binaries_file"

if [ "$declared" -eq 0 ]; then
    echo "$PROGNAME: la lista declarada está vacía: el contrato de tooling no está declarado y el preflight falla, no pasa en silencio" >&2
    exit 2
fi

if [ "$missing_count" -gt 0 ]; then
    if [ "$runner_given" -eq 1 ]; then
        # Una línea exacta por ausente, en el orden de la declaración: el
        # mensaje dice QUÉ falta y EN QUÉ runner falta (regla 16).
        printf '%s\n' "$missing" | while IFS= read -r entry; do
            [ -n "$entry" ] || continue
            printf 'falta %s en el runner %s\n' "$entry" "$runner_name" >&2
        done
    else
        # Invocación heredada: el comportamiento genérico previo, intacto.
        echo "$PROGNAME: contrato de tooling incumplido en el runner que ejecuta este job ($missing_count de $declared ausentes):" >&2
        printf '%s\n' "$missing" >&2
        echo "$PROGNAME: el fix aterriza en la definición del servicio, nunca en una instalación manual del host (regla 17)" >&2
    fi
    exit 1
fi

if [ "$runner_given" -eq 1 ]; then
    echo "$PROGNAME: contrato de tooling verificado en el runner $runner_name ($declared binarios presentes)"
else
    echo "$PROGNAME: contrato de tooling verificado en este runner ($declared binarios presentes)"
fi
exit 0
