#!/bin/sh
# branch-name.sh — generador y validador determinista de nombres de rama
# `<tipo>/<N>-<slug>` (skill ci-pattern, HR-35; hueco G11 del inventario de
# activos portable en `references/asset-inventory.md` de esta skill).
#
# Regla: el nombre de rama se GENERA, nunca se prescribe de memoria. Un
# nombre que el gate del repo rechaza se regenera con este script, no se
# reescribe a mano.
#
# Uso:
#   branch-name.sh new --issue N --type TIPO (--slug TEXTO | --slug-from TEXTO)
#                      [--pattern REGEX] [--gate-script RUTA] [--max-len N]
#       Construye `<tipo>/<N>-<slug>` con el slug normalizado, lo valida
#       contra la regex y lo imprime por stdout. Exit 1 si no casa.
#   branch-name.sh validate <rama> [--pattern REGEX] [--gate-script RUTA]
#                      [--max-len N]
#       Valida un nombre de rama existente (pre-delegación). Exit 1 e
#       imprime la forma esperada si no casa.
#   branch-name.sh slugify <texto>
#       Slug determinista: minúsculas, sin acentos, solo [a-z0-9-].
#   branch-name.sh self-test
#       Batería de casos (válido, sin número, mayúsculas, acentos, largo,
#       issue no numérica) ejecutable sin el repo destino. Exit 0 solo si
#       todos los casos pasan.
#
# La regex es un PARÁMETRO del consumer (parámetro 1 de
# `assets/parameters.md`): por defecto se usa la del origen verificado
# (`ardelperal/APAP_WEB`, `scripts/check_branch_name.py`). En la adopción
# pásela con `--pattern` o deje que este script la extraiga del gate del
# destino con `--gate-script <ruta a check_branch_name.py>`.

set -u

PROGNAME=${0##*/}
DEFAULT_PATTERN='^(?:(chore|feat|fix|perf|refactor|docs|ci|test)/[0-9]+-[a-z0-9-]+|archive/.+|main)$'
DEFAULT_MAX_LEN=100
EXAMPLE='feat/441-branch-name-gate'
# Literal salto de línea (así, no con $(printf), que recorta los \n
# finales) y retorno de carro, para rechazar nombres no representables
# como rama antes del match por líneas de grep.
NL='
'
CR=$(printf '\r')

fail() { printf '%s: %s\n' "$PROGNAME" "$1" >&2; exit 1; }

usage() {
  sed -n '3,30p' "$0" | sed 's/^# \{0,1\}//'
}

# Slug determinista: minúsculas, sin acentos, solo [a-z0-9-], guiones
# colapsados, sin guión inicial ni final.
slugify() {
  printf '%s' "$1" \
    | tr '[:upper:]' '[:lower:]' \
    | sed \
        -e 's/á/a/g' -e 's/à/a/g' -e 's/ä/a/g' -e 's/â/a/g' -e 's/ã/a/g' \
        -e 's/Á/a/g' -e 's/À/a/g' -e 's/Ä/a/g' -e 's/Â/a/g' -e 's/Ã/a/g' \
        -e 's/é/e/g' -e 's/è/e/g' -e 's/ë/e/g' -e 's/ê/e/g' \
        -e 's/É/e/g' -e 's/È/e/g' -e 's/Ë/e/g' -e 's/Ê/e/g' \
        -e 's/í/i/g' -e 's/ì/i/g' -e 's/ï/i/g' -e 's/î/i/g' \
        -e 's/Í/i/g' -e 's/Ì/i/g' -e 's/Ï/i/g' -e 's/Î/i/g' \
        -e 's/ó/o/g' -e 's/ò/o/g' -e 's/ö/o/g' -e 's/ô/o/g' -e 's/õ/o/g' \
        -e 's/Ó/o/g' -e 's/Ò/o/g' -e 's/Ö/o/g' -e 's/Ô/o/g' -e 's/Õ/o/g' \
        -e 's/ú/u/g' -e 's/ù/u/g' -e 's/ü/u/g' -e 's/û/u/g' \
        -e 's/Ú/u/g' -e 's/Ù/u/g' -e 's/Ü/u/g' -e 's/Û/u/g' \
        -e 's/ñ/n/g' -e 's/Ñ/n/g' -e 's/ç/c/g' -e 's/Ç/c/g' -e 's/ß/ss/g' \
    | tr -cs 'a-z0-9' '-' \
    | sed -e 's/^-//' -e 's/-$//'
}

# Extrae la primera cadena entrecomillada tras un `re.compile(` del gate
# de referencia (`check_branch_name.py`). Best-effort: si el gate del
# destino cambia de forma, use `--pattern`.
extract_pattern() {
  [ -f "$1" ] || fail "gate-script no encontrado: $1"
  grep -A3 're\.compile(' "$1" \
    | grep -m1 -oE '"[^"]+"' \
    | sed -e 's/^"//' -e 's/"$//'
}

# Python → ERE: los grupos sin captura `(?:` no existen en ERE POSIX
# (grep -E); se convierten a grupos de captura. Limitación documentada:
# la regex del consumer debe quedar ERE-compatible tras esta conversión
# (la del origen verificado lo es).
to_ere() {
  printf '%s' "$1" | sed 's/(?:/(/g'
}

validate_name() {
  # $1 = nombre de rama, $2 = regex, $3 = longitud máxima
  name=$1
  pattern=$(to_ere "$2")
  maxlen=$3
  if [ -z "$name" ]; then
    printf '%s: nombre de rama vacío\n' "$PROGNAME" >&2
    return 1
  fi
  # Un salto de línea o retorno de carro no es representable como rama:
  # grep evalúa línea a línea, así que el rechazo va ANTES del match
  # (si no, un sufijo tipo '\nmain' casaría la segunda línea).
  case $name in
    *"$NL"*)
      printf '%s: nombre de rama con salto de línea no representable como rama\n' "$PROGNAME" >&2
      return 1 ;;
    *"$CR"*)
      printf '%s: nombre de rama con retorno de carro no representable como rama\n' "$PROGNAME" >&2
      return 1 ;;
  esac
  len=$(printf '%s' "$name" | wc -c)
  if [ "$len" -gt "$maxlen" ]; then
    printf '%s: nombre de rama demasiado largo (%d > %d): %s\n' \
      "$PROGNAME" "$len" "$maxlen" "$name" >&2
    return 1
  fi
  if printf '%s' "$name" | grep -qE -- "$pattern"; then
    return 0
  fi
  printf '%s: nombre de rama inválido: %s\n' "$PROGNAME" "$name" >&2
  printf 'Expected shape: <tipo>/<N>-<slug> (tipos y excepciones según la regex del repo).\n' >&2
  printf 'Pattern: %s\n' "$pattern" >&2
  printf 'Example: %s\n' "$EXAMPLE" >&2
  return 1
}

pattern="" gate="" maxlen=$DEFAULT_MAX_LEN
issue="" type="" slug="" slug_from=""
positional1=""

# Opciones en cualquier orden (antes o después de los posicionales).
# Las desconocidas fallan ruidoso; los argumentos posicionales se
# recogen en positional1 (cada subcomando valida cuántos acepta).
parse_opts() {
  positional1=""
  while [ $# -gt 0 ]; do
    case $1 in
      --pattern)     [ $# -ge 2 ] || fail "$1 necesita un valor"; pattern=$2; shift 2 ;;
      --gate-script) [ $# -ge 2 ] || fail "$1 necesita un valor"; gate=$2; shift 2 ;;
      --max-len)
        [ $# -ge 2 ] || fail "$1 necesita un valor"
        case $2 in
          ''|*[!0-9]*) fail "--max-len debe ser un entero positivo: obtuvo '$2'" ;;
        esac
        [ "$2" -gt 0 ] || fail "--max-len debe ser un entero positivo: obtuvo '$2'"
        maxlen=$2; shift 2 ;;
      --issue)       [ $# -ge 2 ] || fail "$1 necesita un valor"; issue=$2; shift 2 ;;
      --type)        [ $# -ge 2 ] || fail "$1 necesita un valor"; type=$2; shift 2 ;;
      --slug)        [ $# -ge 2 ] || fail "$1 necesita un valor"; slug=$2; shift 2 ;;
      --slug-from)   [ $# -ge 2 ] || fail "$1 necesita un valor"; slug_from=$2; shift 2 ;;
      -h|--help)     usage; exit 0 ;;
      -*)            fail "opción desconocida: $1" ;;
      *)
        [ -z "$positional1" ] || fail "demasiados argumentos posicionales: $1"
        positional1=$1; shift ;;
    esac
  done
}

resolve_pattern() {
  if [ -n "$pattern" ] && [ -n "$gate" ]; then
    fail "elija --pattern o --gate-script, no ambos"
  fi
  if [ -n "$gate" ]; then
    pattern=$(extract_pattern "$gate")
    [ -n "$pattern" ] || fail "no se pudo extraer la regex de $gate; pase --pattern"
  fi
  [ -n "$pattern" ] || pattern=$DEFAULT_PATTERN
}

cmd_new() {
  parse_opts "$@"
  [ -z "$positional1" ] || fail "new no acepta argumentos posicionales: $positional1"
  resolve_pattern
  [ -n "$issue" ] || fail "falta --issue (la rama lleva el número de issue; HR-35)"
  case $issue in
    *[!0-9]*|'') fail "--issue debe ser numérico: obtuvo '$issue'" ;;
  esac
  [ -n "$type" ] || fail "falta --type (chore|feat|fix|perf|refactor|docs|ci|test...)"
  if [ -n "$slug" ] && [ -n "$slug_from" ]; then
    fail "elija --slug o --slug-from, no ambos"
  fi
  if [ -n "$slug_from" ]; then
    slug=$(slugify "$slug_from")
  else
    slug=$(slugify "$slug")
  fi
  [ -n "$slug" ] || fail "el slug quedó vacío tras la normalización"
  name="$type/$issue-$slug"
  if validate_name "$name" "$pattern" "$maxlen"; then
    printf '%s\n' "$name"
  else
    exit 1
  fi
}

cmd_validate() {
  parse_opts "$@"
  [ -n "$positional1" ] || fail "validate necesita el nombre de rama"
  name=$positional1
  resolve_pattern
  if validate_name "$name" "$pattern" "$maxlen"; then
    printf '%s: OK (%s casa con la regex)\n' "$PROGNAME" "$name"
  else
    exit 1
  fi
}

cmd_slugify() {
  [ $# -ge 1 ] || fail "slugify necesita el texto"
  slugify "$1"
  printf '\n'
}

cmd_self_test() {
  passes=0
  failures=0
  ok()   { passes=$((passes + 1));   printf 'PASS %s\n' "$1"; }
  ko()   { failures=$((failures + 1)); printf 'FAIL %s\n' "$1"; }
  ptrn=$DEFAULT_PATTERN
  mlen=$DEFAULT_MAX_LEN

  # 1. Slug determinista: acentos, mayúsculas y puntuación.
  got=$(slugify 'Código de Despliegue: Verificación ÁÉÍÓÚ Ñ')
  if [ "$got" = "codigo-de-despliegue-verificacion-aeiou-n" ]; then
    ok "slugify normaliza acentos y puntuación ($got)"
  else
    ko "slugify: esperado 'codigo-de-despliegue-verificacion-aeiou-n', obtenido '$got'"
  fi

  # 2. Slug idempotente.
  once=$(slugify 'Árbol Ñandú')
  if [ "$once" = "$(slugify "$once")" ]; then
    ok "slugify es idempotente ($once)"
  else
    ko "slugify no es idempotente: '$once' -> '$(slugify "$once")'"
  fi

  # 3. Nombre válido pasa la validación.
  if validate_name "feat/441-branch-name-gate" "$ptrn" "$mlen" 2>/dev/null; then
    ok "válido: feat/441-branch-name-gate"
  else
    ko "válido rechazado: feat/441-branch-name-gate"
  fi

  # 4. Rama 'main' y 'archive/...' válidas según la regex del origen.
  if validate_name "main" "$ptrn" "$mlen" 2>/dev/null \
     && validate_name "archive/2026-09-auditoria" "$ptrn" "$mlen" 2>/dev/null; then
    ok "válido: main y archive/* (según la regex del origen)"
  else
    ko "main o archive/* rechazadas con la regex por defecto"
  fi

  # 5. Falta el número de issue: el fallo del 5× de la sesión.
  if validate_name "chore/runner-deploy-spof" "$ptrn" "$mlen" 2>/dev/null; then
    ko "sin número de issue aceptado: chore/runner-deploy-spof"
  else
    ok "sin número de issue rechazado: chore/runner-deploy-spof"
  fi

  # 6. Mayúsculas en el slug.
  if validate_name "feat/441-Branch-Name" "$ptrn" "$mlen" 2>/dev/null; then
    ko "mayúsculas aceptadas: feat/441-Branch-Name"
  else
    ok "mayúsculas rechazadas: feat/441-Branch-Name"
  fi

  # 7. Caracteres fuera de [a-z0-9-].
  if validate_name "feat/441-branch_name" "$ptrn" "$mlen" 2>/dev/null; then
    ko "guión bajo aceptado: feat/441-branch_name"
  else
    ok "guión bajo rechazado: feat/441-branch_name"
  fi

  # 8. Nombre con acentos crudo se rechaza; su slug generado pasa.
  if validate_name "feat/441-código" "$ptrn" "$mlen" 2>/dev/null; then
    ko "acréntimo crudo aceptado: feat/441-código"
  else
    ok "acréntimo crudo rechazado: feat/441-código"
  fi
  got=$(cmd_new --issue 441 --type feat --slug-from 'Código' 2>/dev/null) \
    && [ "$got" = "feat/441-codigo" ] \
    && ok "generación desde texto con acentos: $got" \
    || ko "generación desde 'Código' no produjo feat/441-codigo (obtuvo '$got')"

  # 9. Demasiado largo se rechaza con --max-len por defecto.
  long_slug=$(slugify "$(awk 'BEGIN { while (i++ < 300) printf "a" }')")
  if validate_name "feat/441-$long_slug" "$ptrn" "$mlen" 2>/dev/null; then
    ko "nombre de 300+ caracteres aceptado"
  else
    ok "nombre demasiado largo rechazado (max-len=$mlen)"
  fi

  # 10. Issue no numérica se rechaza en `new`.
  if (cmd_new --issue abc --type feat --slug x) >/dev/null 2>&1; then
    ko "issue no numérica aceptada: abc"
  else
    ok "issue no numérica rechazada: abc"
  fi

  # 11. Caso real de la sesión: el generador produce lo que el gate pide.
  got=$(cmd_new --issue 7 --type chore --slug 'Runner-Deploy SPOF' 2>/dev/null)
  if [ "$got" = "chore/7-runner-deploy-spof" ] \
     && validate_name "$got" "$ptrn" "$mlen" 2>/dev/null; then
    ok "caso de sesión regenerado: chore/runner-deploy-spof -> $got"
  else
    ko "regeneración del caso de sesión falló (obtuvo '$got')"
  fi

  # --- #181: opciones en cualquier orden y nombres no representables ---

  # 12. --pattern DESPUÉS del nombre (orden documentado): debe aplicar la
  #     regex del consumer, no la por defecto.
  if (cmd_validate 'feat/1-x' --pattern '^release/.+$') >/dev/null 2>&1; then
    ko "--pattern tras el nombre ignorado: feat/1-x pasa con ^release/.+$"
  else
    ok "--pattern tras el nombre respetado (no-casante rechazado)"
  fi

  # 13. --pattern ANTES del nombre: debe validar la rama, no la opción.
  if (cmd_validate --pattern '^feat/.+$' 'feat/1-x') >/dev/null 2>&1; then
    ok "--pattern antes del nombre valida la rama"
  else
    ko "--pattern antes del nombre no valida la rama"
  fi

  # 14. --max-len se aplica en validate.
  if (cmd_validate 'feat/1-x' --max-len 5) >/dev/null 2>&1; then
    ko "--max-len ignorado en validate: nombre de 8 caracteres pasa con max-len 5"
  else
    ok "--max-len respetado en validate"
  fi

  # 15. --max-len no entero positivo: rechazo con código != 0 y mensaje.
  rc=0
  err=$( (cmd_new --issue 1 --type feat --slug x --max-len abc) 2>&1 >/dev/null ) || rc=$?
  if [ "$rc" -ne 0 ] && [ -n "$err" ]; then
    ok "--max-len no entero rechazado con mensaje: $err"
  elif [ "$rc" -eq 0 ]; then
    ko "--max-len 'abc' aceptado (control de longitud desactivado)"
  else
    ko "--max-len 'abc' rechazado pero sin mensaje explícito"
  fi

  # 16. Salto de línea: grep evalúa línea a línea; el nombre completo
  #     debe rechazarse, no solo la línea que casa.
  nlname=$(printf 'garbage name\nmain')
  if validate_name "$nlname" "$ptrn" "$mlen" 2>/dev/null; then
    ko "nombre con salto de línea aceptado (grep casó 'main')"
  else
    ok "nombre con salto de línea rechazado"
  fi

  # 17. Retorno de carro con patrón sin anclar (la regex por defecto
  #     lo rechaza de facto; con patrón sin anclar no debía pasar).
  if validate_name "$(printf 'main\rx')" 'main' "$mlen" 2>/dev/null; then
    ko "nombre con retorno de carro aceptado con patrón sin anclar"
  else
    ok "nombre con retorno de carro rechazado"
  fi

  # 18. Nombre vacío.
  if validate_name '' "$ptrn" "$mlen" 2>/dev/null; then
    ko "nombre vacío aceptado"
  else
    ok "nombre vacío rechazado"
  fi

  printf 'self-test: %d pasan, %d fallan\n' "$passes" "$failures"
  [ "$failures" -eq 0 ]
}

case ${1:-} in
  new)        shift; cmd_new "$@" ;;
  validate)   shift; cmd_validate "$@" ;;
  slugify)    shift; cmd_slugify "$@" ;;
  self-test)  cmd_self_test ;;
  -h|--help)  usage; exit 0 ;;
  '')         usage >&2; exit 2 ;;
  *)          usage >&2; fail "subcomando desconocido: $1" ;;
esac
