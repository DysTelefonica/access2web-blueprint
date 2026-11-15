#!/usr/bin/env bash
# gh-api-retry — reintenta una llamada de la CLI `gh` que falla (#1210, #1214).
#
# Una sola regla, reutilizable por todos los gates que consultan la API:
#
#   - Se reintenta un fallo de la llamada: `gh` sale distinto de 0.
#   - Se reintenta una respuesta malformada: el cuerpo no es JSON.
#   - Una respuesta válida se devuelve tal cual y NUNCA se reintenta: una
#     lista vacía puede significar «sin PR» o «cero issues abiertas», y
#     reintentarla rompería el fail-closed del gate que la usa.
#   - Agotados los intentos: sale con 1 y el último error por stderr.
#
# Los cuatro casos viven en `casos.tsv`, la MISMA tabla que recorre la suite
# del cliente Python (`runner-controls/tests/test_github_rest_retry.py`): las dos
# implementaciones no pueden divergir sin que salte un test.
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

# ¿El cuerpo es JSON? Sin parser y sin dependencias: la API responde objetos
# o arrays, asi que un cuerpo vacio o que no empieza por `{` o `[` es una
# respuesta malformada (C7 de #1220).
_es_json() {
  case "$1" in
    "{"*|"["*) return 0 ;;
    *) return 1 ;;
  esac
}

intento=1
while true; do
  : > "$err_file"
  if salida="$(gh "$@" 2>"$err_file")"; then
    if _es_json "$salida"; then
      cat "$err_file" >&2
      printf '%s' "$salida"
      exit 0
    fi
    # La llamada respondio pero el cuerpo no es JSON: se trata como
    # malformada y se reintenta (#1220).
    ultima="$salida"$'\n'"$(cat "$err_file")"
  else
    # gh deja el cuerpo del error (p. ej. el JSON de un 404) en stdout.
    ultima="$salida"$'\n'"$(cat "$err_file")"
  fi
  if [[ "$intento" -ge "$INTENTOS" ]]; then
    echo "ERROR: la llamada a la API falló o devolvió un cuerpo que no es JSON tras $intento intentos: $ultima" >&2
    exit 1
  fi
  echo "AVISO: la llamada a la API falló o devolvió un cuerpo que no es JSON (intento $intento de $INTENTOS); reintento en ${ESPERA_S}s: $ultima" >&2
  sleep "$ESPERA_S"
  intento=$((intento + 1))
done
