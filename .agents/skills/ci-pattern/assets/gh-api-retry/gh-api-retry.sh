#!/usr/bin/env bash
# gh-api-retry — reintenta una llamada de la CLI `gh` que falla (#1210, #1214, #1592).
#
# Una sola regla, reutilizable por todos los gates que consultan la API:
#
#   - Se reintenta lo transitorio: un error de red (`gh` sale distinto de 0
#     sin código HTTP), un 5xx, un 429 y un cuerpo que no es JSON.
#   - Un 4xx distinto de 429 NO se reintenta: sale con 1 al primer intento.
#   - Backoff exponencial: la primera espera es GH_API_RETRY_ESPERA_S y cada
#     reintento la dobla (por defecto, 4 intentos: 2, 4 y 8 s).
#   - Una respuesta válida se devuelve tal cual y NUNCA se reintenta: una
#     lista vacía puede significar «sin PR» o «cero issues abiertas», y
#     reintentarla rompería el fail-closed del gate que la usa.
#   - Agotados los intentos: sale con 1 y el último error por stderr.
#
# Los casos viven en `casos.tsv`, la MISMA tabla que recorre la suite del
# cliente Python (`runner-controls/tests/test_github_rest_retry.py`): las dos
# implementaciones no pueden divergir sin que salte un test.
#
# Uso:
#   gh-api-retry.sh <args de gh...>
#
# Entorno:
#   CI_PATTERN_GH          gh a ejecutar (#1609, #1632): si está fijada, ella; una
#                          ruta .py corre con python3. Si no, `gh` del PATH.
#   GH_API_RETRY_INTENTOS  intentos totales (4 por defecto; 1 desactiva el reintento)
#   GH_API_RETRY_ESPERA_S  primera espera en segundos (2 por defecto); cada
#                          reintento la dobla. `0` = sin esperas (tests).
#
# La salida de `gh` va a stdout y su stderr a stderr, sin mezclarse: un
# aviso de `gh` no puede corromper el JSON que el gate va a parsear.
set -uo pipefail

INTENTOS="${GH_API_RETRY_INTENTOS:-4}"
ESPERA_S="${GH_API_RETRY_ESPERA_S:-2}"

if ! [[ "$INTENTOS" =~ ^[0-9]+$ ]] || [[ "$INTENTOS" -lt 1 ]]; then
  echo "ERROR: GH_API_RETRY_INTENTOS debe ser un entero >= 1 (valor: $INTENTOS)" >&2
  exit 2
fi
# Una espera que no es un número no negativo vuelve a la de por defecto.
if ! [[ "$ESPERA_S" =~ ^[0-9]+([.][0-9]+)?$ ]]; then
  ESPERA_S=2
fi

err_file="$(mktemp)"
trap 'rm -f "$err_file"' EXIT

# El gh que se ejecuta: el mismo contrato que los assets de #1609. CI_PATTERN_GH,
# si está fijada (una ruta .py corre con python3); si no, `gh` del PATH. Así un
# gh real delante en el PATH no se invoca cuando el test fija el falso.
_gh() {
  if [[ -n "${CI_PATTERN_GH:-}" && "$CI_PATTERN_GH" == *.py ]]; then
    python3 "$CI_PATTERN_GH" "$@"
  else
    "${CI_PATTERN_GH:-gh}" "$@"
  fi
}

# ¿El cuerpo es JSON? Sin parser y sin dependencias: la API responde objetos
# o arrays, asi que un cuerpo vacio o que no empieza por `{` o `[` es una
# respuesta malformada (C7 de #1220).
_es_json() {
  case "$1" in
    "{"*|"["*) return 0 ;;
    *) return 1 ;;
  esac
}

# Código HTTP que `gh` dice en stderr («(HTTP 404)»); vacío si fue error de red.
_http_status() {
  sed -n 's/.*(HTTP \([0-9][0-9][0-9]\)).*/\1/p' "$1" | tail -n 1
}

intento=1
while true; do
  : > "$err_file"
  if salida="$(_gh "$@" 2>"$err_file")"; then
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
    status="$(_http_status "$err_file")"
    # Un 4xx distinto de 429 no es transitorio: no se reintenta (#1592).
    if [[ "$status" == 4* && "$status" != 429 ]]; then
      echo "ERROR: la llamada a la API respondió HTTP $status (un 4xx no se reintenta): $ultima" >&2
      exit 1
    fi
  fi
  if [[ "$intento" -ge "$INTENTOS" ]]; then
    echo "ERROR: la llamada a la API falló o devolvió un cuerpo que no es JSON tras $intento intentos: $ultima" >&2
    exit 1
  fi
  echo "AVISO: la llamada a la API falló o devolvió un cuerpo que no es JSON (intento $intento de $INTENTOS); reintento en ${ESPERA_S}s: $ultima" >&2
  sleep "$ESPERA_S"
  ESPERA_S="$(awk -v s="$ESPERA_S" 'BEGIN { printf "%g", s * 2 }')"
  intento=$((intento + 1))
done
