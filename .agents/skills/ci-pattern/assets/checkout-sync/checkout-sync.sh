#!/bin/sh
# ci-pattern:agent-guards v1 (checkout-sync) — control local del checkout atrasado
# (DysTelefonica/team-skills#914). Local, sin red: compara HEAD con refs/remotes/origin/<base>
# tal como quedaron en la ultima sincronizacion. Si origin/<base> no es ancestro de HEAD,
# el checkout esta atrasado, y el mensaje dice el comando segun el caso:
#   - rama de flota (skill-fleet/*) o HEAD desacoplado: sincroniza <base>;
#   - estas en <base>: sincroniza tu rama (git pull --ff-only);
#   - rama de trabajo: integra origin/<base> en tu rama (git merge).
# Codigos: 0 sincronizado; 1 atrasado o rama equivocada; 2 no puede medir
# (sin origin, sin sincronizar, ultima sincronizacion vieja). Nunca 0 si no mide.
#
# Uso: checkout-sync.sh [--repo DIR] [--base RAMA] [--max-age-hours N]
set -u

repo="."
base=""
max_h=24
while [ $# -gt 0 ]; do
  case "$1" in
    --repo) repo="${2:?--repo requiere un valor}"; shift 2 ;;
    --base) base="${2:?--base requiere un valor}"; shift 2 ;;
    --max-age-hours) max_h="${2:?--max-age-hours requiere un valor}"; shift 2 ;;
    *) echo "OPCION no reconocida: $1 (ver cabecera del script)." >&2; exit 2 ;;
  esac
done
case "$max_h" in ''|*[!0-9]*) echo "ABORT: --max-age-hours debe ser un entero." >&2; exit 2 ;; esac

git -C "$repo" rev-parse --show-toplevel >/dev/null 2>&1 || {
  echo "NO PUEDO MEDIR: «$repo» no es un checkout git." >&2; exit 2; }
gitdir=$(git -C "$repo" rev-parse --absolute-git-dir 2>/dev/null) || {
  echo "NO PUEDO MEDIR: sin directorio git en «$repo»." >&2; exit 2; }

if [ -z "$base" ]; then
  base=$(git -C "$repo" config --get ci-pattern.base-branch 2>/dev/null)
fi
if [ -z "$base" ]; then
  base=$(git -C "$repo" symbolic-ref -q --short refs/remotes/origin/HEAD 2>/dev/null | sed 's|^origin/||')
fi
[ -n "$base" ] || base=main

git -C "$repo" remote get-url origin >/dev/null 2>&1 || {
  echo "NO PUEDO MEDIR: sin remoto «origin»: no hay contra qué comparar la rama base «$base»." >&2
  exit 2; }
rref="refs/remotes/origin/$base"
git -C "$repo" rev-parse --verify -q "$rref" >/dev/null || {
  echo "NO PUEDO MEDIR: no existe origin/$base en tu copia local. Ejecuta: git fetch origin" >&2
  exit 2; }

# Sin red: solo se puede medir hasta la ultima sincronizacion. Un fetch viejo daria verde falso.
fh="$gitdir/FETCH_HEAD"
[ -f "$fh" ] || {
  echo "NO PUEDO MEDIR: este checkout nunca se ha sincronizado con origin. Ejecuta: git fetch origin" >&2
  exit 2; }
now=$(date +%s)
mt=$(stat -c %Y "$fh" 2>/dev/null || stat -f %m "$fh" 2>/dev/null) || {
  echo "NO PUEDO MEDIR: no leo la fecha de $fh." >&2; exit 2; }
age_s=$((now - mt))
if [ "$age_s" -gt $((max_h * 3600)) ]; then
  echo "NO PUEDO MEDIR: la ultima sincronizacion fue hace $((age_s / 3600)) h (limite $max_h h) y la comparacion usaria datos viejos. Ejecuta: git fetch origin" >&2
  exit 2
fi

# AGENTS.md es la gobernanza que se lee al arrancar: se nombra en el mensaje.
if git -C "$repo" diff --quiet "$rref" HEAD -- AGENTS.md; then
  gov="AGENTS.md es igual en HEAD y en origin/$base"
else
  gov="AGENTS.md difiere entre HEAD y origin/$base"
fi

branch=$(git -C "$repo" symbolic-ref -q --short HEAD 2>/dev/null)
sync_base="git fetch origin && git switch $base && git pull --ff-only"
if [ -z "$branch" ]; then
  echo "RAMA EQUIVOCADA: HEAD desacoplado, no estas en una rama. Sincroniza $base: $sync_base" >&2
  exit 1
fi
case "$branch" in
  skill-fleet/*)
    echo "RAMA EQUIVOCADA: estas en la rama de flota «$branch», no en tu rama de trabajo. Sincroniza $base: $sync_base" >&2
    exit 1 ;;
esac

if git -C "$repo" merge-base --is-ancestor "$rref" HEAD; then
  echo "OK: HEAD contiene origin/$base ($gov). Medido con la sincronizacion de hace $((age_s / 60)) min."
  exit 0
fi

if [ "$branch" = "$base" ]; then
  echo "ATRASADO: estas en $base y origin/$base tiene commits que tu copia no tiene ($gov). Para ponerte al día, sincroniza tu rama: git fetch origin && git pull --ff-only (si no avanza hay divergencia: no fuerces, resuélvelo a mano)." >&2
  exit 1
fi
echo "ATRASADO: tu rama de trabajo «$branch» no contiene origin/$base ($gov). Para ponerte al día, integra $base en tu rama: git fetch origin && git merge origin/$base" >&2
exit 1
