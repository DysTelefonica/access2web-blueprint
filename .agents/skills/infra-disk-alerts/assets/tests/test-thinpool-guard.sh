#!/usr/bin/env bash
# Suite thinpool-guard (#1430): el aviso de thin pool, contra stubs.
#
# `lvs`, `curl`, `python3`, `logger` y `hostname` son stubs en PATH: la suite
# no toca LVM ni la red. Cada notificación queda registrada en $CALLS como
# `telegram <primera línea>` o `gmail <asunto>`.
#
# Las reglas: alerta por encima del umbral (datos o metadatos), una sola vez
# dentro de la ventana de repetición, de nuevo al vencer la ventana; un aviso
# de recuperación al bajar del umbral; Gmail solo cuando Telegram falla; si no
# sale ningún aviso, el marcador no se crea y la corrida falla para reintentar.
set -uo pipefail

ASSET="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SCRIPT="$ASSET/thinpool-guard.sh"
PASS=0
FAIL=0
FAILURES=()

assert_eq() {
  if [[ "$2" == "$3" ]]; then
    PASS=$((PASS + 1))
  else
    FAIL=$((FAIL + 1))
    FAILURES+=("$1: esperaba '$2', obtuvo '$3'")
  fi
}

# Sin directorio temporal, los stubs irían a "/bin": se aborta antes.
TMP="$(mktemp -d)" || { echo "thinpool-guard: mktemp falló" >&2; exit 1; }
[[ -n "$TMP" && -d "$TMP" ]] || { echo "thinpool-guard: sin directorio temporal" >&2; exit 1; }
trap 'rm -rf "$TMP"' EXIT
# Forma POSIX: una ruta `C:/...` metería dos puntos en PATH (Git Bash).
TMP="$(cd "$TMP" && pwd)"
mkdir -p "$TMP/bin"

cat > "$TMP/bin/lvs" <<'STUB'
#!/usr/bin/env bash
case " $* " in
  *" --noheadings "*) printf '  %s\n' "$FAKE_LVS" ;;
  *) printf '  LV   LSize  Data%%\n  data 141g   %s\n' "$FAKE_LVS" ;;
esac
STUB
cat > "$TMP/bin/curl" <<'STUB'
#!/usr/bin/env bash
cat > /dev/null
[[ "${FAKE_TG_FAIL:-0}" == 1 ]] && exit 22
text=""
while [[ $# -gt 0 ]]; do
  [[ "$1" == "--data-urlencode" && "$2" == text=* ]] && text="${2#text=}"
  shift
done
printf 'telegram %s\n' "$(printf '%s\n' "$text" | head -1)" >> "$CALLS"
STUB
cat > "$TMP/bin/python3" <<'STUB'
#!/usr/bin/env bash
cat > /dev/null
[[ "${FAKE_MAIL_FAIL:-0}" == 1 ]] && exit 1
printf 'gmail %s\n' "$MAIL_SUBJECT" >> "$CALLS"
STUB
printf '#!/usr/bin/env bash\nexit 0\n' > "$TMP/bin/logger"
printf '#!/usr/bin/env bash\necho testhost\n' > "$TMP/bin/hostname"
chmod +x "$TMP/bin/"*

printf 'TELEGRAM_BOT_TOKEN=123:abc\nTELEGRAM_CHAT_ID=42\n' > "$TMP/tg.env"
printf 'abcdabcdabcdabcd\n' > "$TMP/pw"

export PATH="$TMP/bin:$PATH"
export CALLS="$TMP/calls"
export STATE_DIR="$TMP/state"
export TELEGRAM_ENV="$TMP/tg.env"
export PASS_FILE="$TMP/pw"
MARKER="$STATE_DIR/alert-pve-data"

run() {
  # $1 = "vg lv data meta"; resto: variables extra para esta corrida.
  local lvs="$1"; shift
  : > "$CALLS"
  env FAKE_LVS="$lvs" "$@" sh "$SCRIPT" > "$TMP/out" 2>&1
  echo $? > "$TMP/rc"
}
calls() { tr '\n' '|' < "$CALLS"; }
marker() { [[ -e "$MARKER" ]] && echo si || echo no; }

# 1. Por encima del umbral: alerta por Telegram y marcador.
run "pve data 87.10 3.20"
assert_eq "alerta_rc" 0 "$(cat "$TMP/rc")"
assert_eq "alerta_canal" "telegram [testhost] ALERTA thin pool pve/data al 87.10%|" "$(calls)"
assert_eq "alerta_marcador" si "$(marker)"

# 2. Sigue alto dentro de la ventana: no repite.
run "pve data 88.00 3.20"
assert_eq "ventana_sin_aviso" "" "$(calls)"

# 3. Ventana vencida: repite.
touch -d '7 hours ago' "$MARKER"
run "pve data 88.00 3.20"
assert_eq "ventana_vencida" "telegram [testhost] ALERTA thin pool pve/data al 88.00%|" "$(calls)"

# 4. Baja del umbral con marcador: recuperación y marcador fuera.
run "pve data 60.00 3.20"
assert_eq "recuperacion" "telegram [testhost] OK thin pool pve/data al 60.00%|" "$(calls)"
assert_eq "recuperacion_marcador" no "$(marker)"

# 5. Por debajo sin marcador: silencio.
run "pve data 60.00 3.20"
assert_eq "normal_sin_aviso" "" "$(calls)"

# 6. Metadatos por encima aunque los datos no: alerta.
run "pve data 10.00 90.00"
assert_eq "metadatos" "telegram [testhost] ALERTA thin pool pve/data al 10.00%|" "$(calls)"
rm -f "$MARKER"

# 7. Telegram falla: sale por Gmail.
run "pve data 87.10 3.20" FAKE_TG_FAIL=1
assert_eq "respaldo_gmail" "gmail [testhost] ALERTA thin pool pve/data al 87.10%|" "$(calls)"
rm -f "$MARKER"

# 8. No sale ningún aviso: sin marcador y corrida fallida, para reintentar.
run "pve data 87.10 3.20" FAKE_TG_FAIL=1 FAKE_MAIL_FAIL=1
assert_eq "sin_canal_marcador" no "$(marker)"
assert_eq "sin_canal_rc" 1 "$([[ "$(cat "$TMP/rc")" != 0 ]] && echo 1 || echo 0)"

# 9. Umbral configurable.
run "pve data 60.00 3.20" THRESHOLD=50
assert_eq "umbral" "telegram [testhost] ALERTA thin pool pve/data al 60.00%|" "$(calls)"
rm -f "$MARKER"

# 10. --test: un aviso de prueba y ningún marcador.
: > "$CALLS"
FAKE_LVS="pve data 60.00 3.20" sh "$SCRIPT" --test > /dev/null 2>&1
assert_eq "prueba" "telegram [testhost] thinpool-guard: prueba|" "$(calls)"
assert_eq "prueba_marcador" no "$(marker)"

echo "thinpool-guard: PASS=$PASS FAIL=$FAIL"
for f in "${FAILURES[@]}"; do echo "  FAIL $f"; done
[[ "$FAIL" -eq 0 ]]
