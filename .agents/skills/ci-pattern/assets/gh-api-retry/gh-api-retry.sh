#!/usr/bin/env bash
# gh-api-retry — reintenta una llamada de la CLI `gh` que falla (#1210, #1214).
#
# Una sola regla, reutilizable por todos los gates que consultan la API:
#
#   - Se reintenta SOLO un fallo de la llamada: `gh` sale distinto de 0.
#   - Una respuesta válida se devuelve tal cual y NUNCA se reintenta: una
#     lista vacía puede significar «sin PR» o «cero issues abiertas», y
#     reintentarla rompería el fail-closed del gate que la usa.
#   - Agotados los intentos: sale con 1 y el último error por stderr.
#
# Uso:
#   gh-api-retry.sh <args de gh...>
#
# Entorno:
#   GH_API_RETRY_INTENTOS  intentos totales (3 por defecto; 1 desactiva el reintento)
#   GH_API_RETRY_ESPERA_S  segundos de espera entre intentos (1 por defecto)
#
# La salida de `gh` va a stdout y su stderr a stderr, sin mezclarse: un
# aviso de `gh` no puede corromper el JSON que el gate va a parsear.
set -uo pipefail

INTENTOS="${GH_API_RETRY_INTENTOS:-3}"
ESPERA_S="${GH_API_RETRY_ESPERA_S:-1}"

if ! [[ "$INTENTOS" =~ ^[0-9]+$ ]] || [[ "$INTENTOS" -lt 1 ]]; then
  echo "ERROR: GH_API_RETRY_INTENTOS debe ser un entero >= 1 (valor: $INTENTOS)" >&2
  exit 2
fi

err_file="$(mktemp)"
trap 'rm -f "$err_file"' EXIT

intento=1
while true; do
  : > "$err_file"
  if salida="$(gh "$@" 2>"$err_file")"; then
    cat "$err_file" >&2
    printf '%s' "$salida"
    exit 0
  fi
  ultima="$(cat "$err_file")"
  if [[ "$intento" -ge "$INTENTOS" ]]; then
    echo "ERROR: la llamada a la API falló tras $intento intentos: $salida$ultima" >&2
    exit 1
  fi
  echo "AVISO: la llamada a la API falló (intento $intento de $INTENTOS); reintento en ${ESPERA_S}s: $salida$ultima" >&2
  sleep "$ESPERA_S"
  intento=$((intento + 1))
done
