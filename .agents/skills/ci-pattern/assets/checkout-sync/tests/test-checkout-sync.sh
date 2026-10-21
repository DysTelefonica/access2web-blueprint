#!/usr/bin/env bash
# test-checkout-sync.sh — comportamiento del control local del checkout
# atrasado (DysTelefonica/team-skills#914). Solo interfaz publica: se invoca
# checkout-sync.sh y se mira el codigo de salida y el texto. Repos desechables
# bajo mktemp; nunca toca este clon.
# Codigos: 0 sincronizado, 1 atrasado o rama equivocada, 2 no puede medir.
set -u

: "${TMPDIR:=${TEMP:-${TMP:-/tmp}}}"
export TMPDIR

THIS_DIR="$(cd -- "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPT="$THIS_DIR/../checkout-sync.sh"

export GIT_CONFIG_GLOBAL=/dev/null GIT_CONFIG_SYSTEM=/dev/null
export GIT_AUTHOR_NAME=t GIT_AUTHOR_EMAIL=t@example.invalid
export GIT_COMMITTER_NAME=t GIT_COMMITTER_EMAIL=t@example.invalid

TSROOT=$(mktemp -d "$TMPDIR/checkout-sync.XXXXXX") || exit 1
trap 'rm -rf -- "$TSROOT"' EXIT

PASS=0
FAIL=0
FAILURES=()
ok()   { PASS=$((PASS+1)); }
fail() { FAIL=$((FAIL+1)); FAILURES+=("$1"); echo "  X $1"; }

# run <desc> <rc-esperado> <repo> [args...]  -> deja OUT y RC
run() {
  local desc="$1" esperado="$2" repo="$3"; shift 3
  OUT=$(sh "$SCRIPT" --repo "$repo" "$@" 2>&1); RC=$?
  if [ "$RC" -eq "$esperado" ]; then ok; else fail "$desc: esperado rc=$esperado, rc=$RC ($OUT)"; fi
}
has() { # <desc> <texto>: la salida contiene el texto
  case "$OUT" in *"$2"*) ok ;; *) fail "$1: falta «$2» en la salida: $OUT" ;; esac
}
hasnot() { # <desc> <texto>: la salida NO contiene el texto
  case "$OUT" in *"$2"*) fail "$1: aparece «$2» y no debe: $OUT" ;; *) ok ;; esac
}

# Origen desechable con main, y un clon de trabajo; «up» es otro colaborador.
o="$TSROOT/origin.git"; seed="$TSROOT/seed"; w="$TSROOT/work"; up="$TSROOT/up"
git init -q --bare -b main "$o"
git init -q -b main "$seed"
echo "# gobierno v1" > "$seed/AGENTS.md"
git -C "$seed" add -A && git -C "$seed" commit -q -m base
git -C "$seed" remote add origin "$o" && git -C "$seed" push -q origin main
git clone -q "$o" "$w"      # sin FETCH_HEAD: ningun fetch todavia
git clone -q "$o" "$up"

# Avanza origin/main desde otro clon; el checkout de trabajo no lo sabe aun.
avanza() {
  echo "# gobierno v$1" > "$up/AGENTS.md"
  git -C "$up" commit -q -am "avance $1" && git -C "$up" push -q origin main
}

# ---------------- sin remoto o sin medir [codigo 2] ----------------
nr="$TSROOT/sin-remoto"; git init -q -b main "$nr"
echo x > "$nr/a" && git -C "$nr" add -A && git -C "$nr" commit -q -m a
run "sin origin: no da verde" 2 "$nr"
has "sin origin: lo dice" "origin"

run "fuera de un checkout git: no da verde" 2 "$TSROOT"
has "fuera de git: lo dice" "git"

run "clon sin ninguna sincronizacion: no da verde" 2 "$w"
has "nunca sincronizado: pide fetch" "git fetch origin"

# ---------------- main al dia [codigo 0] ----------------
git -C "$w" fetch -q origin
run "main al dia: verde" 0 "$w"
has "verde nombra la base" "origin/main"

# ---------------- rama base atrasada: sincroniza tu rama [codigo 1] ----------------
avanza 2 && git -C "$w" fetch -q origin
run "main atrasada: rojo" 1 "$w"
has "base atrasada: sincroniza tu rama" "sincroniza tu rama"
has "base atrasada: comando exacto" "git pull --ff-only"
has "nombra AGENTS.md" "AGENTS.md"
has "AGENTS.md difiere" "difiere"
hasnot "base atrasada: no manda integrar" "git merge origin/main"

git -C "$w" pull -q --ff-only
echo local > "$w/local.txt" && git -C "$w" add -A && git -C "$w" commit -q -m "local"
run "main adelantada a origin: verde" 0 "$w"

# Divergencia en la base: pull --ff-only no basta; se dice sin forzar.
avanza 3 && git -C "$w" fetch -q origin
run "main divergente: rojo" 1 "$w"
has "divergente: no forzar" "no fuerces"
git -C "$w" merge -q origin/main
run "main integrada con origin: verde" 0 "$w"

# ---------------- rama de trabajo atrasada: integra main [codigo 1] ----------------
git -C "$w" switch -q -c feat/914-x
avanza 4 && git -C "$w" fetch -q origin
run "rama de trabajo atrasada: rojo" 1 "$w"
has "rama de trabajo: integra main" "integra main en tu rama"
has "rama de trabajo: comando exacto" "git merge origin/main"
hasnot "rama de trabajo: no dice pull --ff-only" "git pull --ff-only"
git -C "$w" merge -q origin/main
run "rama de trabajo tras integrar main: verde" 0 "$w"

# ---------------- rama de flota o HEAD desacoplado [codigo 1] ----------------
git -C "$w" switch -q -c skill-fleet/cadete
run "rama de flota: rojo aunque este al dia" 1 "$w"
has "flota: sincroniza main" "rama de flota"
has "flota: vuelve a la base" "git switch main"
git -C "$w" switch -q feat/914-x

git -C "$w" checkout -q --detach
run "HEAD desacoplado: rojo" 1 "$w"
has "desacoplado: lo dice" "desacoplado"
git -C "$w" switch -q feat/914-x

# ---------------- ultima sincronizacion vieja [codigo 2] ----------------
git -C "$w" fetch -q origin
run "sincronizacion reciente: verde" 0 "$w"
touch -d '3 days ago' "$w/.git/FETCH_HEAD"
run "ultimo fetch viejo: no da verde" 2 "$w"
has "fetch viejo: pide fetch" "git fetch origin"
run "umbral configurable: verde si cabe" 0 "$w" --max-age-hours 200

# ---------------- subcomando del CLI: delega y conserva el codigo ----------------
cli="$THIS_DIR/../../bin/ci-pattern"
OUT=$(python3 "$cli" checkout-sync --repo "$w" --max-age-hours 200 2>&1); RC=$?
if [ "$RC" -eq 0 ]; then ok; else fail "CLI checkout-sync sincronizado: rc=$RC ($OUT)"; fi
OUT=$(python3 "$cli" checkout-sync --repo "$w" 2>&1); RC=$?
if [ "$RC" -eq 2 ]; then ok; else fail "CLI checkout-sync con fetch viejo: rc=$RC ($OUT)"; fi

# ---------------- resumen ----------------
echo
echo "==== resumen checkout-sync (#914) ===="
echo "pasados: $PASS"
echo "fallidos: $FAIL"
if [ "$FAIL" -gt 0 ]; then
  printf '  - %s\n' "${FAILURES[@]}"
  exit 1
fi
exit 0
