#!/usr/bin/env bash
# test_tooling_preflight.sh — suite del preflight de tooling (#165).
#
# Contrato verificado (runner-skill reglas 16/17):
#   - lista declarada completa y presente en PATH  -> exit 0;
#   - binario ausente                              -> exit != 0 y NOMBRE al
#     usuario (todos los ausentes, no solo el primero);
#   - lista declarada vacía (solo comentarios y
#     líneas en blanco)                            -> exit != 0: un contrato
#     no declarado falla, nunca pasa en silencio;
#   - fichero de lista inexistente o ilegible      -> exit != 0.
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

# stub PATH: sólo un dir con stubs + /usr/bin + /bin (sed del propio script).
STUB_DIR="$(mktemp -d)"
trap 'rm -rf "$STUB_DIR"' EXIT
make_stub() {
    printf '#!/bin/sh\nexit 0\n' > "$STUB_DIR/$1"
    chmod +x "$STUB_DIR/$1"
}
ISOLATED_PATH="$STUB_DIR:/usr/bin:/bin"

# run_case <nombre> <exit-esperado: 0|nonzero> <fichero-lista> [patrón-que-debe-aparecer]
run_case() {
    local name=$1 want=$2 list=$3 pattern=${4:-}
    local out rc
    out=$(PATH="$ISOLATED_PATH" "$PROG" --binaries "$list" 2>&1)
    rc=$?
    if [ "$want" = "0" ]; then
        if [ "$rc" -eq 0 ]; then ok "$name"; else ko "$name (exit $rc: $out)"; fi
        return
    fi
    if [ "$rc" -eq 0 ]; then ko "$name (exit 0, esperaba != 0: $out)"; return; fi
    if [ -n "$pattern" ] && ! printf '%s' "$out" | grep -q "$pattern"; then
        ko "$name (exit $rc pero la salida no nombra '$pattern': $out)"
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

# --- Caso 3: binario ausente -> != 0 y lo nombra.
LIST_MISSING="$STUB_DIR/list-missing.txt"
printf 'present-tool-a\nno-existe-tool-xyz\n' > "$LIST_MISSING"
run_case "binario ausente -> exit != 0 nombrandolo" nonzero "$LIST_MISSING" "no-existe-tool-xyz"

# --- Caso 4: varios ausentes -> TODOS nombrados, no solo el primero.
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
if PATH="$ISOLATED_PATH" "$PROG" >/dev/null 2>&1; then
    ko "sin --binaries -> exit != 0 (exit 0)"
else
    ok "sin --binaries -> exit != 0"
fi

echo
echo "tooling-preflight: $PASS ok, $FAIL fail"
[ "$FAIL" -eq 0 ]
