#!/usr/bin/env bash
# Suite gh-api-retry: el reintento de la llamada a la API (#1210, #1214).
#
# El helper reintenta SOLO un fallo de la llamada (gh sale distinto de 0).
# Una respuesta válida se devuelve tal cual: una lista vacía puede
# significar «sin PR» y reintentarla rompería el fail-closed del gate
# que la usa. Agotados los intentos, sale con 1 y el último error por
# stderr.
set -uo pipefail

SCRIPT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/gh-api-retry.sh"
PASS=0
FAIL=0
FAILURES=()
TMP=""

assert_exit() {
  local desc="$1" expected="$2" actual="$3"
  if [[ "$expected" == "$actual" ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$desc: esperaba exit $expected, obtuvo $actual")
  fi
}

assert_contains() {
  if [[ "$3" == *"$2"* ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$1: esperaba contener '$2'")
  fi
}

assert_not_contains() {
  if [[ "$3" != *"$2"* ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$1: no esperaba contener '$2'")
  fi
}

assert_eq() {
  if [[ "$2" == "$3" ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$1: esperaba '$2', obtuvo '$3'")
  fi
}

cleanup() { [[ -n "$TMP" ]] && rm -rf "$TMP"; }
trap cleanup EXIT

# gh falso: falla las primeras <fails> llamadas y después responde <json>.
# Cada invocación se registra en calls para poder contarlas.
make_gh() {
  local fails="$1" body="$2"
  cat > "$TMP/gh" <<'GH'
#!/usr/bin/env bash
set -u
printf 'x\n' >> "$FAKE_GH_CALLS"
n="$(wc -l < "$FAKE_GH_CALLS" | tr -d ' ')"
if [[ "$n" -le "$FAKE_GH_FAILS" ]]; then
  echo "unexpected end of JSON input"
  exit 1
fi
printf '%s\n' "$FAKE_GH_BODY"
GH
  chmod +x "$TMP/gh"
  export FAKE_GH_FAILS="$fails" FAKE_GH_BODY="$body"
  : > "$TMP/calls"
  export FAKE_GH_CALLS="$TMP/calls"
}

run_helper() {
  # Variables globales a proposito: el harness las lee despues.
  out="$(env PATH="$TMP:$PATH" bash "$SCRIPT" api "repos/acme/repo/pulls" 2>"$TMP/err")" && rc=0 || rc=$?
  err="$(cat "$TMP/err")"
  CALLS="$(wc -l < "$TMP/calls" | tr -d ' ')"
}

TMP="$(mktemp -d)"

# 1. La llamada funciona a la primera: la salida pasa intacta y no se
# reintenta.
make_gh 0 '[{"number": 1}]'
run_helper
assert_exit "llamada correcta sale 0" 0 "$rc"
assert_eq "la salida pasa intacta" '[{"number": 1}]' "$out"
assert_eq "una llamada correcta no se reintenta" 1 "$CALLS"

# 2. Fallo transitorio: la segunda llamada responde y la salida es la del
# exito, no la del fallo.
make_gh 1 '[{"number": 2}]'
run_helper
assert_exit "fallo transitorio sale 0" 0 "$rc"
assert_eq "la salida es la del intento correcto" '[{"number": 2}]' "$out"
assert_eq "el fallo transitorio consume un reintento" 2 "$CALLS"
assert_contains "avisa del reintento" "AVISO" "$err"

# 3. Todos los intentos fallan: fail closed con el ultimo error.
make_gh 5 'nunca'
run_helper
assert_exit "todos los intentos fallando salen 1" 1 "$rc"
assert_eq "se agotan los intentos" 3 "$CALLS"
assert_contains "el ultimo error queda dicho" "unexpected end of JSON input" "$err"

# 4. Una respuesta VALIDA vacia no se reintenta: puede significar «sin PR».
make_gh 0 ''
run_helper
assert_exit "respuesta vacia sale 0" 0 "$rc"
assert_eq "la respuesta vacia se devuelve tal cual" "" "$out"
assert_eq "la respuesta vacia no se reintenta" 1 "$CALLS"
assert_not_contains "la respuesta vacia no avisa de reintento" "AVISO" "$err"

# 5. Con un solo intento, un fallo no se reintenta.
make_gh 3 'nunca'
OUT="$(env PATH="$TMP:$PATH" GH_API_RETRY_INTENTOS=1 bash "$SCRIPT" api "repos/acme/repo/pulls" 2>"$TMP/err")" && RC=0 || RC=$?
assert_exit "un solo intento sale 1" 1 "$RC"
assert_eq "un solo intento no reintenta" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_not_contains "un solo intento no deja el error en stdout" "unexpected end" "$OUT"

echo ""
if [[ "$FAIL" -gt 0 ]]; then
  echo "FAIL: $FAIL fallos de $((PASS + FAIL)) aserciones (test-gh-api-retry)"
  for f in "${FAILURES[@]}"; do
    echo "  - $f"
  done
  exit 1
fi
echo "PASS: $PASS aserciones, 0 fallos (test-gh-api-retry)"
