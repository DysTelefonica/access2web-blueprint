#!/bin/sh
# thinpool-guard: alert by email when an LVM thin pool fills up.
#
# A full thin pool takes every guest on the host down at once, and the CI
# containers here are overprovisioned (their disks add up to more than the
# pool). This checks data and metadata usage of every thin pool and mails an
# alert above THRESHOLD, repeating every REPEAT_HOURS while it stays high, and
# one recovery mail when it drops back below.
#
# Alerts go to Telegram first (bot token and chat id in TELEGRAM_ENV, the
# same file vps-disk-guard reads); if Telegram fails or is not configured,
# they fall back to Gmail SMTP with the app password in PASS_FILE. Both files
# are root-only (0600) and no secret reaches argv or the log.
#
# Usage: thinpool-guard          check (run from cron)
#        thinpool-guard --test   send a test mail and exit
set -eu

THRESHOLD="${THRESHOLD:-85}"
REPEAT_HOURS="${REPEAT_HOURS:-6}"
MAIL_USER="${MAIL_USER:-ardelperal@gmail.com}"
MAIL_TO="${MAIL_TO:-ardelperal@gmail.com}"
PASS_FILE="${PASS_FILE:-/root/.gmail-app-password}"
TELEGRAM_ENV="${TELEGRAM_ENV:-/root/.vps-disk-guard-telegram.env}"
STATE_DIR="${STATE_DIR:-/var/lib/thinpool-guard}"
HOST="$(hostname)"

log() { logger -t thinpool-guard "$*"; echo "$(date -Is) $*"; }

send_mail() {
  # $1 subject, body on stdin.
  MAIL_SUBJECT="$1" MAIL_USER="$MAIL_USER" MAIL_TO="$MAIL_TO" PASS_FILE="$PASS_FILE" \
  python3 -c '
import os, smtplib, sys
from email.message import EmailMessage
m = EmailMessage()
m["Subject"] = os.environ["MAIL_SUBJECT"]
m["From"] = os.environ["MAIL_USER"]
m["To"] = os.environ["MAIL_TO"]
m.set_content(sys.stdin.read())
pw = open(os.environ["PASS_FILE"]).read().replace(" ", "").strip()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as s:
    s.login(os.environ["MAIL_USER"], pw)
    s.send_message(m)
'
}

send_telegram() {
  # $1 text. Fails (non-zero) when Telegram is not configured or refuses.
  [ -f "$TELEGRAM_ENV" ] || return 1
  (
    # shellcheck disable=SC1090
    . "$TELEGRAM_ENV"
    [ -n "${TELEGRAM_BOT_TOKEN:-}" ] && [ -n "${TELEGRAM_CHAT_ID:-}" ] || exit 1
    # The token goes to curl through a config on stdin, never argv or the log.
    printf 'url = "https://api.telegram.org/bot%s/sendMessage"\n' "$TELEGRAM_BOT_TOKEN" \
      | curl --silent --fail --max-time 20 --config - -o /dev/null \
          --data-urlencode "chat_id=$TELEGRAM_CHAT_ID" --data-urlencode "text=$1"
  )
}

notify() {
  # $1 subject, body on stdin: Telegram first, Gmail as fallback.
  body="$(cat)"
  if send_telegram "$1

$body"; then echo telegram; return 0; fi
  printf '%s\n' "$body" | send_mail "$1" && echo gmail
}

mkdir -p "$STATE_DIR"

if [ "${1:-}" = "--test" ]; then
  ch=$(printf 'Prueba del aviso de thin pool en %s.\n\n%s\n' "$HOST" "$(lvs -o vg_name,lv_name,data_percent,metadata_percent -S 'lv_attr=~^t' 2>&1)" \
    | notify "[$HOST] thinpool-guard: prueba")
  log "test alert sent via $ch"
  exit 0
fi

lvs --noheadings -o vg_name,lv_name,data_percent,metadata_percent -S 'lv_attr=~^t' |
while read -r vg lv data meta; do
  marker="$STATE_DIR/alert-$vg-$lv"
  high=$(awk -v d="$data" -v m="$meta" -v t="$THRESHOLD" 'BEGIN { print (d+0 >= t || m+0 >= t) ? 1 : 0 }')
  if [ "$high" = 1 ]; then
    if [ -e "$marker" ] && [ -z "$(find "$marker" -mmin "+$((REPEAT_HOURS * 60))")" ]; then
      log "$vg/$lv data=$data% meta=$meta% over ${THRESHOLD}%, alert already sent within ${REPEAT_HOURS}h"
      continue
    fi
    ch=$({ printf 'El thin pool %s/%s de %s supera el %s%%.\n\n' "$vg" "$lv" "$HOST" "$THRESHOLD"
      printf 'Datos: %s%%  Metadatos: %s%%\n\n' "$data" "$meta"
      printf 'Si llega al 100%%, todos los CT/VM del host se quedan sin disco.\n\n'
      lvs -o lv_name,lv_size,data_percent,origin -S "vg_name=$vg" 2>&1
    } | notify "[$HOST] ALERTA thin pool $vg/$lv al $data%")
    touch "$marker"
    log "ALERT $vg/$lv data=$data% meta=$meta% sent via $ch"
  elif [ -e "$marker" ]; then
    ch=$(printf 'El thin pool %s/%s de %s vuelve a estar por debajo del %s%%.\n\nDatos: %s%%  Metadatos: %s%%\n' \
      "$vg" "$lv" "$HOST" "$THRESHOLD" "$data" "$meta" \
      | notify "[$HOST] OK thin pool $vg/$lv al $data%")
    rm -f "$marker"
    log "RECOVERED $vg/$lv data=$data% meta=$meta% sent via $ch"
  fi
done
