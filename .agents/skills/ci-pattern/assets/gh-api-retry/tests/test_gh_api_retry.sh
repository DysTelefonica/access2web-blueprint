#!/usr/bin/env bash
# Suite gh-api-retry: la tabla de casos es el contrato (#1210, #1220).
#
# Esta suite recorre `casos.tsv`, la MISMA tabla que recorre la suite del
# cliente Python (`runner-controls/tests/test_github_rest_retry.py`). Las dos
# implementaciones del reintento se prueban contra las mismas filas: si una
# diverge de la tabla, su suite falla.
#
# La regla: se reintenta un fallo de llamada y un JSON malformado; una
# respuesta válida vacía NO se reintenta; agotados los intentos, se falla
# cerrado.
set -uo pipefail

ASSET="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SCRIPT="$ASSET/gh-api-retry.sh"
TABLA="$ASSET/casos.tsv"
PASS=0
FAIL=0
FAILURES=()
TMP=""

assert_eq() {
  if [[ "$2" == "$3" ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$1: esperaba '$2', obtuvo '$3'")
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

cleanup() { [[ -n "$TMP" ]] && rm -rf "$TMP"; }
trap cleanup EXIT

# gh falso: falla las primeras <fallos> llamadas segun el modo de fallo y,
# despues, responde segun el modo de respuesta. Cada invocacion se registra
# en calls para poder contar las consultas.
make_gh() {
  cat > "$TMP/gh" <<'GH'
#!/usr/bin/env bash
set -u
printf 'x\n' >> "$FAKE_GH_CALLS"
n="$(wc -l < "$FAKE_GH_CALLS" | tr -d ' ')"
if [[ "$n" -le "$FAKE_GH_FAILS" ]]; then
  case "$FAKE_GH_FAIL_MODE" in
    error) echo "unexpected end of JSON input"; exit 1 ;;
    json_malformado) printf '<html>502 Bad Gateway</html>\n'; exit 0 ;;
  esac
fi
case "$FAKE_GH_RESPONSE" in
  vacia) printf '[]\n' ;;
  *) printf '[{"number": 1, "merged": true, "merged_at": "2026-01-01T00:00:00Z"}]\n' ;;
esac
GH
  chmod +x "$TMP/gh"
}

TMP="$(mktemp -d)"

# --- La tabla, fila por fila ---------------------------------------------
while IFS=$'\t' read -r caso falla fallos intentos respuesta resultado; do
  [[ -z "${caso:-}" || "$caso" == \#* ]] && continue
  : > "$TMP/calls"
  export FAKE_GH_CALLS="$TMP/calls"
  export FAKE_GH_FAILS="$fallos"
  export FAKE_GH_FAIL_MODE="$falla"
  export FAKE_GH_RESPONSE="$respuesta"
  make_gh

  # stdin a /dev/null: el command substitution no puede comerse la tabla.
  out="$(env PATH="$TMP:$PATH" GH_API_RETRY_ESPERA_S=0 bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
  calls="$(wc -l < "$TMP/calls" | tr -d ' ')"
  err="$(cat "$TMP/err")"

  assert_eq "$caso: numero de llamadas" "$intentos" "$calls"
  case "$resultado" in
    ok)
      assert_eq "$caso: sale 0" 0 "$rc"
      assert_eq "$caso: devuelve la respuesta valida" '[{"number": 1, "merged": true, "merged_at": "2026-01-01T00:00:00Z"}]' "$out"
      ;;
    vacia)
      assert_eq "$caso: sale 0" 0 "$rc"
      assert_eq "$caso: devuelve la respuesta vacia tal cual" "[]" "$out"
      assert_eq "$caso: no avisa de reintento" 0 "$(printf '%s' "$err" | grep -c "AVISO" || true)"
      ;;
    falla_cerrado)
      assert_eq "$caso: sale 1 (fallo cerrado)" 1 "$rc"
      assert_eq "$caso: dice cuantos intentos hizo" 1 "$(printf '%s' "$err" | grep -c "3 intentos" || true)"
      ;;
  esac
done < "$TABLA"

# --- Casos concretos del helper (no covered por la tabla) ------------------
# Un intento y nada mas: un fallo no se reintenta.
: > "$TMP/calls"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=3 FAKE_GH_FAIL_MODE=error FAKE_GH_RESPONSE=ok
make_gh
out="$(env PATH="$TMP:$PATH" GH_API_RETRY_INTENTOS=1 GH_API_RETRY_ESPERA_S=0 bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "un solo intento sale 1" 1 "$rc"
assert_eq "un solo intento no reintenta" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_not_contains "un solo intento no deja el error en stdout" "unexpected end" "$out"

# Agotados los intentos, el ultimo error queda dicho en el mensaje final.
: > "$TMP/calls"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=3 FAKE_GH_FAIL_MODE=error FAKE_GH_RESPONSE=ok
make_gh
out="$(env PATH="$TMP:$PATH" GH_API_RETRY_ESPERA_S=0 bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "todos los intentos fallando salen 1" 1 "$rc"
assert_contains "el ultimo error queda dicho" "unexpected end of JSON input" "$(cat "$TMP/err")"

# Un 404 de gh: su cuerpo JSON (stdout) tambien llega al mensaje final.
mkdir -p "$TMP/nf"
cat > "$TMP/nf/gh" <<'GH'
#!/usr/bin/env bash
echo '{"message":"Not Found"}'
echo 'gh: Not Found (HTTP 404)' >&2
exit 1
GH
chmod +x "$TMP/nf/gh"
out="$(env PATH="$TMP/nf:$PATH" GH_API_RETRY_INTENTOS=1 GH_API_RETRY_ESPERA_S=0 bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "un 404 sale 1" 1 "$rc"
assert_contains "un 404 conserva el cuerpo de gh" '{"message":"Not Found"}' "$(cat "$TMP/err")"

# Un valor de intentos invalido se rechaza con codigo 2 (no se reintenta a
# la infinita, ni se acepta un cero que haria fallar la primera llamada).
out="$(env PATH="$TMP:$PATH" GH_API_RETRY_INTENTOS=0 bash "$SCRIPT" api "repos/o/r" 2>&1 </dev/null)" && rc=0 || rc=$?
assert_eq "intentos invalidos salen 2" 2 "$rc"

echo ""
if [[ "$FAIL" -gt 0 ]]; then
  echo "FAIL: $FAIL fallos de $((PASS + FAIL)) aserciones (test-gh-api-retry)"
  for f in "${FAILURES[@]}"; do
    echo "  - $f"
  done
  exit 1
fi
echo "PASS: $PASS aserciones, 0 fallos (test-gh-api-retry)"
