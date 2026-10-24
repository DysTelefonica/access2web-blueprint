#!/usr/bin/env bash
# test_tooling_preflight.sh — suite del preflight de tooling (#165; contrato
# con identidad de runner, TS#1113 clase M).
#
# Contrato verificado (runner-skill reglas 16/17):
#   - lista declarada completa y presente en PATH  -> exit 0;
#   - binario ausente                              -> exit != 0 y NOMBRE al
#     usuario (todos los ausentes, no solo el primero);
#   - binario ausente                              -> exit != 0 y stderr
#     contiene el mensaje exacto
#     `falta <binario> en el runner <runner>`;
#   - lista declarada vacía (solo comentarios y
#     líneas en blanco)                            -> exit != 0: un contrato
#     no declarado falla, nunca pasa en silencio;
#   - fichero de lista inexistente o ilegible      -> exit != 0;
#   - sin --binaries                                -> exit != 0;
#   - invocación heredada sin --runner (lista
#     presente)                                     -> exit 0: la ausencia de
#     `--runner` NO es una invocación inválida.
#
# Identidad del runner: la opción `--runner <nombre>` declara en qué runner
# se ejecuta el job, porque "falta el binario" sin saber dónde no dice
# nada accionable: el fix aterriza en la definición del servicio del runner
# concreto, no en un host cualquiera. El nombre de runner del test es
# determinista (`runner-de-prueba`): el mensaje exacto que se exige depende
# de él y no puede depender del entorno. La opción es opt-in: la invocación
# heredada (`--binaries FICHERO` a secas) sigue siendo válida y conserva su
# comportamiento observable (caso 9).
#
# Aislamiento: el PATH del test se reduce a un directorio con stubs + los
# del sistema mínimos para que el propio preflight funcione; los nombres de
# binario de prueba no colisionan con nada real. Sin red, stdlib.
#
# No compara el exit code exacto de las rutas de fallo (1 ausencias /
# 2 entrada inválida): la issue exige «exit != 0 nombrando al ausente».

set -u

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_ROOT="$(cd -- "$THIS_DIR/../../.." && pwd)"
PROG="$SKILL_ROOT/assets/tooling-preflight/tooling-preflight.sh"

PASS=0
FAIL=0
FAILURES=()

ok() { PASS=$((PASS + 1)); echo "ok - $1"; }
ko() { FAIL=$((FAIL + 1)); FAILURES+=("$1"); echo "FAIL - $1"; }

# Identidad del runner declarada en cada invocación (valor determinista).
RUNNER_ID="runner-de-prueba"

# stub PATH: sólo un dir con stubs + /usr/bin + /bin (sed del propio script).
STUB_DIR="$(mktemp -d)"
OUT_DIR="$(mktemp -d)"
trap 'rm -rf "$STUB_DIR" "$OUT_DIR"' EXIT
make_stub() {
    printf '#!/bin/sh\nexit 0\n' > "$STUB_DIR/$1"
    chmod +x "$STUB_DIR/$1"
}
ISOLATED_PATH="$STUB_DIR:/usr/bin:/bin"

# run_case <nombre> <exit-esperado: 0|nonzero> <fichero-lista> [patrón-en-stderr]
# Invoca el preflight declarando la identidad del runner y comprueba el
# código de salida y, si se pide, el patrón exacto en stderr (stdout y stderr
# se leen por separado: el mensaje de fallo es de stderr).
run_case() {
    local name=$1 want=$2 list=$3 pattern=${4:-}
    local rc out err
    PATH="$ISOLATED_PATH" "$PROG" --binaries "$list" --runner "$RUNNER_ID" \
        > "$OUT_DIR/stdout" 2> "$OUT_DIR/stderr"
    rc=$?
    out=$(cat "$OUT_DIR/stdout")
    err=$(cat "$OUT_DIR/stderr")
    if [ "$want" = "0" ]; then
        if [ "$rc" -eq 0 ]; then ok "$name"; else ko "$name (exit $rc; stdout: $out; stderr: $err)"; fi
        return
    fi
    if [ "$rc" -eq 0 ]; then ko "$name (exit 0, esperaba != 0; stderr: $err)"; return; fi
    if [ -n "$pattern" ] && ! printf '%s' "$err" | grep -qF "$pattern"; then
        ko "$name (exit $rc pero stderr no contiene '$pattern'; stdout: $out; stderr: $err)"
        return
    fi
    ok "$name"
}

# --- Caso 1: todo presente -> 0.
make_stub present-tool-a
make_stub present-tool-b
LIST_OK="$STUB_DIR/list-ok.txt"
printf 'present-tool-a\npresent-tool-b\n' > "$LIST_OK"
run_case "lista completa presente -> exit 0" 0 "$LIST_OK"

# --- Caso 2: comentarios y líneas en blanco del dato versionado se ignoran.
LIST_COMMENTS="$STUB_DIR/list-comments.txt"
printf '# contrato de tooling del job (dato versionado)\n\npresent-tool-a\n' > "$LIST_COMMENTS"
run_case "comentarios y líneas en blanco ignorados -> exit 0" 0 "$LIST_COMMENTS"

# --- Caso 3: binario ausente -> != 0 y mensaje EXACTO con el runner.
# El patrón es el mensaje completo: dice QUÉ falta y EN QUÉ runner falta.
LIST_MISSING="$STUB_DIR/list-missing.txt"
printf 'present-tool-a\nno-existe-tool-xyz\n' > "$LIST_MISSING"
run_case "binario ausente -> exit != 0 con mensaje exacto del runner" nonzero "$LIST_MISSING" \
    "falta no-existe-tool-xyz en el runner runner-de-prueba"

# --- Caso 4: varios ausentes -> TODOS nombrados, no solo el primero.
# Patrón genérico a propósito: aquí se comprueba que se nombran los DOS
# ausentes de una lista con dos ausentes; el mensaje exacto con identidad de
# runner lo fija el caso 3 (lista con un solo ausente).
LIST_SEVERAL="$STUB_DIR/list-several.txt"
printf 'no-existe-tool-xyz\nno-existe-tool-abc\n' > "$LIST_SEVERAL"
run_case "varios ausentes: primero nombrado" nonzero "$LIST_SEVERAL" "no-existe-tool-xyz"
run_case "varios ausentes: segundo tambien nombrado" nonzero "$LIST_SEVERAL" "no-existe-tool-abc"

# --- Caso 5: lista vacía (solo comentarios/blancos) -> fallo.
LIST_EMPTY="$STUB_DIR/list-empty.txt"
printf '# sin declaracion todavia\n\n' > "$LIST_EMPTY"
run_case "lista vacia -> exit != 0" nonzero "$LIST_EMPTY"

# --- Caso 6: fichero de lista inexistente -> fallo.
run_case "lista inexistente -> exit != 0" nonzero "$STUB_DIR/no-such-list.txt"

# --- Caso 7: fichero de lista ilegible -> fallo.
LIST_UNREADABLE="$STUB_DIR/list-unreadable.txt"
printf 'present-tool-a\n' > "$LIST_UNREADABLE"
chmod 000 "$LIST_UNREADABLE"
run_case "lista ilegible -> exit != 0" nonzero "$LIST_UNREADABLE"

# --- Caso 8: sin --binaries -> fallo (contrato no declarado no pasa en silencio).
# Runner declarado, lista no: la invocación es válida en su forma, lo que
# falta es el contrato.
if PATH="$ISOLATED_PATH" "$PROG" --runner "$RUNNER_ID" >/dev/null 2>&1; then
    ko "sin --binaries -> exit != 0 (exit 0)"
else
    ok "sin --binaries -> exit != 0"
fi

# --- Caso 8b: invocación sin ninguna opción -> fallo.
if PATH="$ISOLATED_PATH" "$PROG" >/dev/null 2>&1; then
    ko "sin opciones -> exit != 0 (exit 0)"
else
    ok "sin opciones -> exit != 0"
fi

# --- Caso 9: compatibilidad — invocación heredada SIN --runner.
# La forma previa (`--binaries FICHERO` a secas) conserva su comportamiento
# observable: con toda la lista presente sale 0. La ausencia de `--runner` no
# es una invocación inválida (eso sería exit 2): `--runner` es la opción
# opt-in que añade el mensaje exacto `falta <cli> en el runner <nombre>`.
PATH="$ISOLATED_PATH" "$PROG" --binaries "$LIST_OK" >/dev/null 2>&1
rc=$?
if [ "$rc" -eq 0 ]; then
    ok "invocación heredada sin --runner (lista presente) -> exit 0"
else
    ko "invocación heredada sin --runner (lista presente) -> exit 0 (exit $rc)"
fi

echo
echo "tooling-preflight: $PASS ok, $FAIL fail"
[ "$FAIL" -eq 0 ]
