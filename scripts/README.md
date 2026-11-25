# scripts/

Scripts de setup y mantenimiento del repo `access2web-blueprint`.

## Scripts

| Script | Propósito | Cuándo correrlo |
|---|---|---|
| `setup-rclone.md` | Guía para configurar rclone con Cloudflare R2 | Una vez por máquina nueva |
| `setup-staging.ps1` | Descarga binarios de R2 → `data/staging/<app>/{frontend,backend,resources,.codegraph-vba}` | Una vez por máquina nueva, o después de `git pull` si los paths cambiaron |
| `sync-to-r2.ps1` | Sube binarios de `data/staging/` → R2 | Cuando regenerás un `.codegraph-vba/`, actualizás desde el staging del repo original, o cambias binarios |
| `migration_progress.py` | Genera `docs/07-migracion/seguimiento/avance.md` y `avance.html` a partir de `fases.json` y de las issues `saneamiento:coexistencia` de GitHub (`gh search issues`); falla con código 2, sin escribir nada, si una issue es de un repositorio sin grupo | Cuando cambia el estado de una issue o de una fase: `python scripts/migration_progress.py` |

## Workflow típico

```powershell
# Setup en una máquina nueva
git clone <repo-url> access2web-blueprint
cd access2web-blueprint
pwsh -File scripts/setup-rclone.md   # seguir los pasos una vez
pwsh -File scripts/setup-staging.ps1  # descargar ~870 MB

# Trabajo normal
# (Dysflow lee de data/staging/<app>/{frontend,backend}/)

# Cuando hay cambios (ej. regeneraste codegraph-vba)
pwsh -File scripts/sync-to-r2.ps1 -App condor    # solo Condor
```

## Pre-requisitos

- **rclone** instalado, con el remote `cloudflare-r2` configurado (ver `setup-rclone.md`).
- **Bucket `access2web-staging-binaries`** creado en R2 (lo crea el `setup-staging.ps1` la primera vez si no existe).
- **PowerShell 7+** (`pwsh`).

## Variables de entorno

Ninguna requerida. Las credenciales viven en `rclone config`, no en env vars.

## Logs

Los scripts escriben logs en `../logs/rclone-*.log`. El directorio `logs/` está en `.gitignore`.
