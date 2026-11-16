#!/usr/bin/env bash
# Suite gh-api-retry: la tabla de casos es el contrato (#1210, #1220, #1592).
#
# Esta suite recorre `casos.tsv`, la MISMA tabla que recorre la suite del
# cliente Python (`runner-controls/tests/test_github_rest_retry.py`). Las dos
# implementaciones del reintento se prueban contra las mismas filas: si una
# diverge de la tabla, su suite falla.
#
# La regla está en la tabla y en `../README.md`. Las esperas las anota un
# `sleep` falso en el PATH, sin dormir.
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

REAL=""

cleanup() { [[ -n "$TMP" ]] && rm -rf "$TMP"; [[ -n "$REAL" ]] && rm -rf "$REAL"; }
trap cleanup EXIT

# Guarda (#1609, #1620): el gh que correria el asset, el primero del PATH que
# fija la llamada ($1), tiene que estar dentro de $TMP. El asset llama a `gh`
# por nombre y no lee CI_PATTERN_GH: la guarda resuelve el mismo PATH.
# Si no, se niega sin lanzar ningun proceso.
guard_gh() {
  local gh=""
  gh="$(PATH="$1"; hash -r; type -P gh)"
  if [[ -z "$gh" || ! -f "$gh" || "$(realpath -- "$gh")" != "$(realpath -- "$TMP")"/* ]]; then
    echo "refused: the gh this test would run (${gh:-none}) is outside the fixture dir $TMP;"       "the test did not run and no call was made (#1620)" >&2
    return 1
  fi
}

# Corre `env <args...>` tras la guarda, con el PATH que fijan los argumentos
# (PATH=...); sale 97 si la guarda se niega.
retry() {
  local path_value="$PATH" arg
  for arg in "$@"; do
    case "$arg" in PATH=*) path_value="${arg#PATH=}" ;; esac
  done
  guard_gh "$path_value" || return 97
  env "$@"
}

# gh falso: falla las primeras <fallos> llamadas segun el modo de fallo y,
# despues, responde segun el modo de respuesta. Cada invocacion se registra
# en calls para poder contar las consultas. Un fallo HTTP imita a `gh api`:
# el cuerpo JSON del error en stdout, `gh: ... (HTTP nnn)` en stderr, sale 1.
make_gh() {
  cat > "$TMP/gh" <<'GH'
#!/usr/bin/env bash
set -u
printf 'x\n' >> "$FAKE_GH_CALLS"
n="$(wc -l < "$FAKE_GH_CALLS" | tr -d ' ')"
if [[ "$n" -le "$FAKE_GH_FAILS" ]]; then
  case "$FAKE_GH_FAIL_MODE" in
    red) echo 'error connecting to api.github.com' >&2; exit 1 ;;
    5xx) echo '{"message":"Service Unavailable"}'; echo 'gh: Service Unavailable (HTTP 503)' >&2; exit 1 ;;
    429) echo '{"message":"rate limited"}'; echo 'gh: rate limited (HTTP 429)' >&2; exit 1 ;;
    4xx) echo '{"message":"Not Found"}'; echo 'gh: Not Found (HTTP 404)' >&2; exit 1 ;;
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

# sleep falso: anota cada espera en $FAKE_SLEEPS y no duerme (#1592).
make_sleep() {
  cat > "$TMP/sleep" <<'SLEEP'
#!/usr/bin/env bash
printf '%s\n' "$1" >> "$FAKE_SLEEPS"
SLEEP
  chmod +x "$TMP/sleep"
}

# Esperas anotadas, separadas por comas; `-` si no hubo ninguna.
esperas() {
  if [[ -s "$FAKE_SLEEPS" ]]; then paste -sd, "$FAKE_SLEEPS"; else echo -; fi
}

TMP="$(mktemp -d)"
unset CI_PATTERN_GH # una variable exportada por el usuario no debe colarse en la suite (#1609, #1632)
make_sleep
export FAKE_SLEEPS="$TMP/sleeps"

# --- La tabla, fila por fila ---------------------------------------------
while IFS=$'\t\r' read -r caso falla fallos intentos respuesta resultado espera_tabla; do
  [[ -z "${caso:-}" || "$caso" == \#* ]] && continue
  : > "$TMP/calls"
  : > "$FAKE_SLEEPS"
  export FAKE_GH_CALLS="$TMP/calls"
  export FAKE_GH_FAILS="$fallos"
  export FAKE_GH_FAIL_MODE="$falla"
  export FAKE_GH_RESPONSE="$respuesta"
  make_gh

  # stdin a /dev/null: el command substitution no puede comerse la tabla.
  # Espera e intentos por defecto (2 s, 4): los fija la tabla; el sleep es falso.
  out="$(retry -u GH_API_RETRY_ESPERA_S -u GH_API_RETRY_INTENTOS PATH="$TMP:$PATH" bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
  calls="$(wc -l < "$TMP/calls" | tr -d ' ')"
  err="$(cat "$TMP/err")"

  assert_eq "$caso: numero de llamadas" "$intentos" "$calls"
  assert_eq "$caso: esperas observadas" "$espera_tabla" "$(esperas)"
  if [[ "$espera_tabla" == "-" ]]; then
    assert_eq "$caso: no avisa de reintento" 0 "$(printf '%s' "$err" | grep -c "AVISO" || true)"
  fi
  case "$resultado" in
    ok)
      assert_eq "$caso: sale 0" 0 "$rc"
      assert_eq "$caso: devuelve la respuesta valida" '[{"number": 1, "merged": true, "merged_at": "2026-01-01T00:00:00Z"}]' "$out"
      ;;
    vacia)
      assert_eq "$caso: sale 0" 0 "$rc"
      assert_eq "$caso: devuelve la respuesta vacia tal cual" "[]" "$out"
      ;;
    falla_cerrado)
      assert_eq "$caso: sale 1 (fallo cerrado)" 1 "$rc"
      assert_eq "$caso: no deja nada en stdout" "" "$out"
      if [[ "$intentos" -gt 1 ]]; then
        assert_eq "$caso: dice cuantos intentos hizo" 1 "$(printf '%s' "$err" | grep -c "tras $intentos intentos" || true)"
      fi
      ;;
  esac
done < "$TABLA"

# --- Casos concretos del helper (no covered por la tabla) ------------------
# Un intento y nada mas: un fallo no se reintenta.
: > "$TMP/calls"; : > "$FAKE_SLEEPS"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=99 FAKE_GH_FAIL_MODE=error FAKE_GH_RESPONSE=ok
make_gh
out="$(retry PATH="$TMP:$PATH" GH_API_RETRY_INTENTOS=1 bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "un solo intento sale 1" 1 "$rc"
assert_eq "un solo intento no reintenta" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_eq "un solo intento no espera" - "$(esperas)"
assert_not_contains "un solo intento no deja el error en stdout" "unexpected end" "$out"

# La espera base del entorno escala el backoff: cada reintento la dobla.
: > "$TMP/calls"; : > "$FAKE_SLEEPS"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=99 FAKE_GH_FAIL_MODE=5xx FAKE_GH_RESPONSE=ok
make_gh
out="$(retry PATH="$TMP:$PATH" GH_API_RETRY_ESPERA_S=0.5 bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "espera base 0.5: sale 1" 1 "$rc"
assert_eq "espera base 0.5: esperas 0.5, 1 y 2" "0.5,1,2" "$(esperas)"
assert_contains "el aviso nombra la espera" "(intento 3 de 4); reintento en 2s" "$(cat "$TMP/err")"

# Una espera base invalida vuelve a la de por defecto (como el cliente Python).
: > "$TMP/calls"; : > "$FAKE_SLEEPS"
out="$(retry PATH="$TMP:$PATH" GH_API_RETRY_ESPERA_S=abc GH_API_RETRY_INTENTOS=2 bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "espera base invalida: usa 2 s" "2" "$(esperas)"

# Un 4xx no se reintenta: un intento, sin espera, y el mensaje lo dice.
: > "$TMP/calls"; : > "$FAKE_SLEEPS"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=99 FAKE_GH_FAIL_MODE=4xx FAKE_GH_RESPONSE=ok
make_gh
out="$(retry PATH="$TMP:$PATH" bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "un 4xx sale 1" 1 "$rc"
assert_eq "un 4xx no se reintenta" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_eq "un 4xx no espera" - "$(esperas)"
assert_contains "un 4xx dice que no se reintenta" "respondió HTTP 404 (un 4xx no se reintenta)" "$(cat "$TMP/err")"
assert_contains "un 4xx conserva el cuerpo de gh" '{"message":"Not Found"}' "$(cat "$TMP/err")"

# Agotados los intentos, el ultimo error queda dicho en el mensaje final.
: > "$TMP/calls"; : > "$FAKE_SLEEPS"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=99 FAKE_GH_FAIL_MODE=error FAKE_GH_RESPONSE=ok
make_gh
out="$(retry PATH="$TMP:$PATH" bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
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
out="$(retry PATH="$TMP/nf:$TMP:$PATH" GH_API_RETRY_INTENTOS=1 bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "un 404 sale 1" 1 "$rc"
assert_contains "un 404 conserva el cuerpo de gh" '{"message":"Not Found"}' "$(cat "$TMP/err")"

# Un valor de intentos invalido se rechaza con codigo 2 (no se reintenta a
# la infinita, ni se acepta un cero que haria fallar la primera llamada).
out="$(retry PATH="$TMP:$PATH" GH_API_RETRY_INTENTOS=0 bash "$SCRIPT" api "repos/o/r" 2>&1 </dev/null)" && rc=0 || rc=$?
assert_eq "intentos invalidos salen 2" 2 "$rc"

# --- CI_PATTERN_GH manda sobre el PATH (#1609, #1632) --------------------
# Un gh "real" delante en el PATH no se invoca: CI_PATTERN_GH apunta al falso.
mkdir -p "$TMP/real"
cat > "$TMP/real/gh" <<'GH'
#!/usr/bin/env bash
echo x >> "$FAKE_REAL_CALLS"
printf '[]\n'
GH
chmod +x "$TMP/real/gh"
: > "$TMP/calls"; : > "$TMP/real_calls"; : > "$FAKE_SLEEPS"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=0 FAKE_GH_FAIL_MODE=error FAKE_GH_RESPONSE=ok FAKE_REAL_CALLS="$TMP/real_calls"
make_gh
out="$(retry PATH="$TMP/real:$PATH" CI_PATTERN_GH="$TMP/gh" bash "$SCRIPT" api "repos/o/r/commits/sha/pulls" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "CI_PATTERN_GH: sale 0" 0 "$rc"
assert_eq "CI_PATTERN_GH: llama al falso una vez" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_eq "CI_PATTERN_GH: el gh real del PATH no se invoca" 0 "$(wc -l < "$TMP/real_calls" | tr -d ' ')"
assert_eq "CI_PATTERN_GH: devuelve la respuesta del falso" '[{"number": 1, "merged": true, "merged_at": "2026-01-01T00:00:00Z"}]' "$out"

# Una ruta .py corre con el Python del asset (el mismo contrato que #1609).
cat > "$TMP/ghpy.py" <<'PY'
import os
with open(os.environ["FAKE_PY_CALLS"], "a", encoding="utf-8") as f:
    f.write("x\n")
print("[]")
PY
: > "$TMP/pycalls"; : > "$TMP/real_calls"
out="$(retry PATH="$TMP/real:$PATH" CI_PATTERN_GH="$TMP/ghpy.py" FAKE_PY_CALLS="$TMP/pycalls" bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "CI_PATTERN_GH .py: sale 0" 0 "$rc"
assert_eq "CI_PATTERN_GH .py: lo corre el Python y no el gh real" 1 "$(wc -l < "$TMP/pycalls" | tr -d ' ')"
assert_eq "CI_PATTERN_GH .py: el gh real no se invoca" 0 "$(wc -l < "$TMP/real_calls" | tr -d ' ')"

# Sin CI_PATTERN_GH, el gh del PATH sigue siendo el que se invoca (las filas de la tabla lo prueban).
# --- #1620: el asset nunca resuelve un gh real primero en el PATH -------
# Un gh "real" de senuelo, primero en el PATH, que anota cualquier llamada.
REAL="$(mktemp -d)"
cat > "$REAL/gh" <<'GH'
#!/bin/sh
echo "$*" >> "$(dirname "$0")/calls"
exit 97
GH
chmod +x "$REAL/gh"
: > "$TMP/calls"
export FAKE_GH_CALLS="$TMP/calls" FAKE_GH_FAILS=0 FAKE_GH_FAIL_MODE=red FAKE_GH_RESPONSE=ok
make_gh
out="$(retry PATH="$REAL:$TMP:$PATH" bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "gh real primero en el PATH: se niega (97)" 97 "$rc"
assert_eq "gh real primero en el PATH: el gh real no recibe llamadas" "" "$(cat "$REAL/calls" 2>/dev/null)"
assert_eq "gh real primero en el PATH: el gh falso no recibe llamadas" "" "$(cat "$TMP/calls")"
assert_contains "gh real primero en el PATH: dice que no hizo ninguna llamada" "no call was made (#1620)" "$(cat "$TMP/err")"

# Con el gh falso primero en el PATH, el asset corre y llama solo al falso.
: > "$TMP/calls"
out="$(retry PATH="$TMP:$REAL:$PATH" bash "$SCRIPT" api "repos/o/r" 2>"$TMP/err" </dev/null)" && rc=0 || rc=$?
assert_eq "gh falso primero en el PATH: sale 0" 0 "$rc"
assert_eq "gh falso primero en el PATH: responde el falso" 1 "$(wc -l < "$TMP/calls" | tr -d ' ')"
assert_eq "gh falso primero en el PATH: el gh real no recibe llamadas" "" "$(cat "$REAL/calls" 2>/dev/null)"

echo ""
if [[ "$FAIL" -gt 0 ]]; then
  echo "FAIL: $FAIL fallos de $((PASS + FAIL)) aserciones (test-gh-api-retry)"
  for f in "${FAILURES[@]}"; do
    echo "  - $f"
  done
  exit 1
fi
echo "PASS: $PASS aserciones, 0 fallos (test-gh-api-retry)"
