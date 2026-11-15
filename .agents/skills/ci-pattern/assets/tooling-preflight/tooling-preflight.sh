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
# conserva su comportamiento observable (mismo texto de ausencia); lo que
# añade `--runner` es el mensaje exacto por CLI ausente.
#
# #1172 (decisión del dueño, 2026-10-24; auditoría de fricción de gates en
# team-skills#1093, comentario `6389037313`): el preflight distingue DOS
# fallos que antes compartían el mismo código de salida.
#
#   - «BINARIO AUSENTE EN EL RUNNER» es un AVISO. El runner no lo tiene y el
#     fix aterriza en la definición de su servicio, no en el PR que se está
#     abriendo: bloquear el PR por eso es fricción pura (la auditoría midió 19
#     job-runs de este gate —8 de fricción y 11 de infraestructura— en la
#     ventana de 48 h). El aviso dice el binario, el runner y la instrucción
#     de instalación, y el preflight sale 0.
#   - «CONTRATO ROTO» —falta `--binaries`, la lista no existe, no es legible o
#     está vacía, `--runner` sin nombre, opción no reconocida— es un BLOQUEO:
#     sale 2. Un contrato sin declarar o ilegible no es un aviso: es un hueco.
#
# Instrucción de instalación: la lista puede declararla por entrada, con el
# comentario `# instalar: <comando>` (o `# install:`). Sin ella, el aviso
# remite a la definición del servicio del runner, que es donde el fix debe
# aterrizar.
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
#   0  el contrato está declarado y es verificable: todos los binarios
#      declarados están presentes, o falta alguno y eso es un AVISO
#      (una línea por ausente: `falta <binario> en el runner <nombre>: <cómo
#      instalarlo>`; sin `--runner`, el resumen con la lista de ausentes);
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
# ausentes (en el orden de la declaración) con su mensaje y, si la declaran,
# su instrucción de instalación. La lista vacía no es un pase: sale 2.
declared=0
missing_count=0
missing=""
missing_msgs=""
missing_hints=""
while IFS= read -r line || [ -n "$line" ]; do
    # Recorte de espacios en blanco extremos (el comentario se separa después,
    # para poder leer la instrucción de instalación que declara).
    raw=$(printf '%s' "$line" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
    case "$raw" in
        '#'*) continue ;;   # línea de comentario suelta
    esac
    name=${raw%%#*}
    name=$(printf '%s' "$name" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
    [ -n "$name" ] || continue
    # Instrucción de instalación declarada por la entrada: `# instalar: <cmd>`
    # (o `# install:`). Sin ella, el aviso remite al servicio del runner.
    hint=""
    case "$raw" in
        *\#*)
            comment=${raw#*#}
            case "$comment" in
                instalar:*|install:*)
                    hint=$(printf '%s' "${comment#*:}" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
                    ;;
            esac
            ;;
    esac
    declared=$((declared + 1))
    if command -v "$name" >/dev/null 2>&1; then
        continue
    fi
    if [ -z "$missing" ]; then
        missing=$name
    else
        missing="$missing
$name"
    fi
    if [ -n "$hint" ]; then
        missing_hints="$missing_hints$name: $hint
"
    fi
    if [ "$runner_given" -eq 1 ]; then
        # Una línea exacta por ausente, en el orden de la declaración: el
        # mensaje dice QUÉ falta, EN QUÉ runner falta y CÓMO se instala
        # (regla 16).
        if [ -n "$hint" ]; then
            msg="falta $name en el runner $runner_name: instalar con: $hint"
        else
            msg="falta $name en el runner $runner_name: instale $name en la definición del servicio del runner $runner_name (no a mano en el host, regla 17)"
        fi
    else
        msg="falta $name: instale $name en la definición del servicio del runner (no a mano en el host, regla 17)"
    fi
    missing_msgs="$missing_msgs$msg
"
    missing_count=$((missing_count + 1))
done < "$binaries_file"

if [ "$declared" -eq 0 ]; then
    echo "$PROGNAME: la lista declarada está vacía: el contrato de tooling no está declarado y el preflight falla, no pasa en silencio" >&2
    exit 2
fi

if [ "$missing_count" -gt 0 ]; then
    # #1172: el binario ausente es un AVISO, no un bloqueo: el fix aterriza
    # en la definición del servicio del runner, no en el PR.
    if [ "$runner_given" -eq 1 ]; then
        printf '%s' "$missing_msgs" | while IFS= read -r msg; do
            [ -n "$msg" ] || continue
            printf '%s\n' "$msg" >&2
            if [ "${GITHUB_ACTIONS:-}" = "true" ]; then
                # Anotación de GitHub Actions: el aviso se ve en el PR sin
                # ponerlo en rojo (va por stderr, donde vive el aviso).
                printf '::warning::%s\n' "$msg" >&2
            fi
        done
    else
        # Invocación heredada: el comportamiento genérico previo, intacto.
        echo "$PROGNAME: AVISO: contrato de tooling incumplido en el runner que ejecuta este job ($missing_count de $declared ausentes):" >&2
        printf '%s\n' "$missing" >&2
        echo "$PROGNAME: el fix aterriza en la definición del servicio, nunca en una instalación manual del host (regla 17)" >&2
        if [ -n "$missing_hints" ]; then
            printf '%s' "$missing_hints" | while IFS= read -r entry; do
                [ -n "$entry" ] || continue
                echo "$PROGNAME: instalar $entry" >&2
            done
        fi
    fi
    # AVISO: el contrato está declarado y es verificable; lo que falta es una
    # herramienta del runner, y eso no bloquea el PR.
    exit 0
fi

if [ "$runner_given" -eq 1 ]; then
    echo "$PROGNAME: contrato de tooling verificado en el runner $runner_name ($declared binarios presentes)"
else
    echo "$PROGNAME: contrato de tooling verificado en este runner ($declared binarios presentes)"
fi
exit 0
