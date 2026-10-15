#!/usr/bin/env bash
# ci-pattern:agent-guards v1 — instalador de guards locales (#420)
# install-agent-guards.sh — materializa los guards de agente (pre-commit y
# pre-push) del patrón en el directorio de hooks REAL del consumer.
#
# Reglas que instalan (ver hooks/pre-commit y hooks/pre-push):
#   - sin commits en la rama protegida (default: main) ni en el checkout
#     principal del consumer (HR-59, DysTelefonica/team-skills#420);
#   - sin push directo a la rama protegida, incluida su eliminación
#     (la rama remota se conserva siempre tras el merge).
#
# Idempotente y no destructivo: un hook existente SIN el marcador
# `ci-pattern:agent-guards` se preserva como `<hook>.ci-pattern-preserved`
# solo con --force; sin --force el instalador aborta nombrando el fichero.
# El desinstalar retira únicamente los ficheros con marcador propio.
#
# Uso:
#   bash install-agent-guards.sh [--repo R] [--protected-branch B]
#                                [--dry-run] [--check] [--uninstall] [--force]
# Salidas:
#   0 operación completa / guards presentes (--check conforme)
#   1 hallazgos (hook ajeno sin --force, --check con guard ausente)
#   2 entorno insalvable (no es git, hooks-path ilegible, backup preexistente)
set -u

SCRIPT_DIR="$(cd -- "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MARKER="ci-pattern:agent-guards v1"
HOOKS=(pre-commit pre-push)

repo="."
protected=""
dry_run=0
check_only=0
uninstall=0
force=0

while [ $# -gt 0 ]; do
  case "$1" in
    --repo) repo="${2:?--repo requiere un valor}"; shift 2 ;;
    --protected-branch) protected="${2:?--protected-branch requiere un valor}"; shift 2 ;;
    --dry-run) dry_run=1; shift ;;
    --check) check_only=1; shift ;;
    --uninstall) uninstall=1; shift ;;
    --force) force=1; shift ;;
    *) echo "OPCION no reconocida: $1 (ver cabecera del script)." >&2; exit 2 ;;
  esac
done

root=$(git -C "$repo" rev-parse --show-toplevel 2>/dev/null) || {
  echo "ABORT: $repo no tiene toplevel de git." >&2
  exit 2
}
hooks_abs=$(git -C "$repo" config --get core.hooksPath 2>/dev/null)
if [ -n "$hooks_abs" ]; then
  case "$hooks_abs" in
    /*) hooks_dir="$hooks_abs" ;;
    *)  hooks_dir="$root/$hooks_abs" ;;
  esac
else
  common=$(git -C "$repo" rev-parse --path-format=absolute --git-common-dir 2>/dev/null) || {
    echo "ABORT: sin --git-common-dir absoluto." >&2; exit 2
  }
  hooks_dir="$common/hooks"
fi

# --check: los dos hooks existen y llevan el marcador propio.
if [ "$check_only" -eq 1 ]; then
  rc=0
  for h in "${HOOKS[@]}"; do
    if [ -f "$hooks_dir/$h" ] && grep -qF "$MARKER" "$hooks_dir/$h"; then
      echo "OK: guard $h instalado en $hooks_dir."
    else
      echo "HALLAZGO: guard $h ausente o sin marcador en $hooks_dir (HR-59 exige ejecutar este instalador al arrancar, #420)." >&2
      rc=1
    fi
  done
  exit "$rc"
fi

if [ "$uninstall" -eq 1 ]; then
  removed=0
  for h in "${HOOKS[@]}"; do
    if [ -f "$hooks_dir/$h" ] && grep -qF "$MARKER" "$hooks_dir/$h"; then
      if [ "$dry_run" -eq 1 ]; then echo "DRY: retiraria $hooks_dir/$h"; else
        rm -f "$hooks_dir/$h" && echo "OK: retirado $hooks_dir/$h" && removed=$((removed+1)); fi
    fi
  done
  exit 0
fi

# Instalación: verificar colisiones antes de escribir nada (todo-o-nada).
for h in "${HOOKS[@]}"; do
  src="$SCRIPT_DIR/hooks/$h"
  [ -f "$src" ] || { echo "ABORT: falta el hook de origen $src." >&2; exit 2; }
  dst="$hooks_dir/$h"
  if [ -f "$dst" ] && ! grep -qF "$MARKER" "$dst"; then
    if [ "$force" -ne 1 ]; then
      echo "HALLAZGO: $dst existe sin el marcador del patron (hook ajeno). Relaje con --force: se preserva como $dst.ci-pattern-preserved (#420)." >&2
      exit 1
    fi
    if [ -e "$dst.ci-pattern-preserved" ]; then
      echo "ABORT: $dst.ci-pattern-preserved ya existe: no se pisa un backup previo. Resuelva a mano (#420, HR-3)." >&2
      exit 2
    fi
    [ "$dry_run" -eq 1 ] || mv "$dst" "$dst.ci-pattern-preserved"
    echo "OK: hook ajeno preservado como $dst.ci-pattern-preserved"
  fi
done

[ -d "$hooks_dir" ] || { [ "$dry_run" -eq 1 ] || mkdir -p "$hooks_dir"; }

if [ -n "$protected" ]; then
  if [ "$dry_run" -eq 1 ]; then
    echo "DRY: git -C $repo config ci-pattern.protected-branch $protected"
  else
    git -C "$repo" config ci-pattern.protected-branch "$protected" || {
      echo "ABORT: no se pudo fijar ci-pattern.protected-branch." >&2; exit 2; }
    echo "OK: rama protegida = $protected"
  fi
fi

for h in "${HOOKS[@]}"; do
  dst="$hooks_dir/$h"
  if [ "$dry_run" -eq 1 ]; then
    echo "DRY: instalarias $dst desde $SCRIPT_DIR/hooks/$h"
    continue
  fi
  cp "$SCRIPT_DIR/hooks/$h" "$dst" || { echo "ABORT: fallo copiando $dst." >&2; exit 2; }
  chmod +x "$dst" || { echo "ABORT: fallo chmod $dst." >&2; exit 2; }
  echo "OK: guard $h instalado en $dst"
done

if [ "$dry_run" -ne 1 ]; then
  echo "OK: guards de agente activos (HR-59, DysTelefonica/team-skills#420)."
fi
exit 0
