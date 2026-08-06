# Script: setup-staging.ps1
# Propósito: descarga los binarios legacy (.accdb) y los índices codegraph-vba desde
# Cloudflare R2 al repo local, en data/staging/<app>/{frontend,backend,resources,.codegraph-vba}.
#
# Las fuentes de verdad:
#   - src/         : staging del repo original (en git)
#   - docs/        : docs del repo original (en git, puede estar outdated)
#   - frontend/    : R2 (este script descarga)
#   - backend/     : R2 (este script descarga)
#   - resources/   : R2 (este script descarga)
#   - .codegraph-vba/ : R2 (este script descarga; regenerable con codegraph-vba index)
#
# Uso:
#   pwsh -File scripts/setup-staging.ps1                # todas las 8 apps
#   pwsh -File scripts/setup-staging.ps1 -App condor     # solo Condor
#   pwsh -File scripts/setup-staging.ps1 -Force         # re-descarga todo
#   pwsh -File scripts/setup-staging.ps1 -DryRun        # ver qué haría
#
# Pre-requisitos:
#   - rclone configurado con el remote "cloudflare-r2" (ver .gitignore vía vault.yml).
#   - Bucket "access2web-staging-binaries" en R2 (ver scripts/setup-rclone.md cuando exista).

param(
    [string]$App = "",                 # '' = todas, 'condor'/'hps'/etc = una sola
    [switch]$Force = $false,
    [switch]$DryRun = $false,
    [string]$RepoRoot = "C:\00repos\codigo\access2web-blueprint"
)

$ErrorActionPreference = "Stop"
$StagingDir = Join-Path $RepoRoot "data\staging"
$Remote = "cloudflare-r2:access2web-staging-binaries"

$Apps = @("condor", "hps", "hps-solicitudes", "brass", "gestion-riesgos", "no-conformidades", "expedientes", "lanzaderas")

if ($App -ne "" -and $Apps -notcontains $App) {
    Write-Host "[ERROR] App '$App' no reconocida. Apps válidas: $($Apps -join ', ')" -ForegroundColor Red
    exit 1
}

function Write-Step($msg) { Write-Host "[SETUP-STAGING] $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "[SKIP] $msg" -ForegroundColor Yellow }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

# 1. Verificar rclone + remote
Write-Step "Verificando rclone..."
$rcloneVersion = (& rclone --version 2>&1 | Select-Object -First 1)
if (-not $rcloneVersion) {
    Write-Err "rclone no está instalado o no está en PATH."
    exit 1
}
Write-OK "rclone: $rcloneVersion"

$remoteList = (& rclone listremotes 2>&1)
if ($remoteList -notcontains "cloudflare-r2") {
    Write-Err "Remote 'cloudflare-r2' no está configurado. Ver scripts/setup-rclone.md."
    exit 1
}
Write-OK "Remote cloudflare-r2 configurado"

# 2. Verificar acceso al bucket
Write-Step "Verificando acceso al bucket $Remote..."
try {
    $bucketList = (& rclone lsd $Remote 2>&1)
    if ($LASTEXITCODE -ne 0) {
        Write-Err "No se puede acceder al bucket. ¿Está creado?"
        exit 1
    }
} catch {
    Write-Err "Error accediendo al bucket: $_"
    exit 1
}
Write-OK "Bucket accesible"

# 3. Para cada app
$targets = if ($App -ne "") { @($App) } else { $Apps }

foreach ($appId in $targets) {
    $appDir = Join-Path $StagingDir $appId
    Write-Step "Procesando $appId"

    if (-not (Test-Path $appDir)) {
        Write-Warn "  $appDir no existe. Esperá que la rama tenga los archivos del repo (src/, docs/, README.md)."
        continue
    }

    foreach ($sub in @("frontend", "backend", "resources", ".codegraph-vba")) {
        $dest = Join-Path $appDir $sub
        if ((Test-Path $dest) -and -not $Force) {
            Write-Skip "  $sub/ ya existe (usa -Force para re-descargar)"
            continue
        }
        if ($DryRun) {
            Write-Host "  [DRY-RUN] rclone copy $Remote/$appId/$sub/ $dest/"
            continue
        }
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        # rclone copy (no sync) — respeta lo local que ya estaba
        & rclone copy "$Remote/$appId/$sub/" "$dest/" --quiet --transfers 4 --checkers 8
        if ($LASTEXITCODE -ne 0) {
            Write-Warn "  $sub/ no se pudo descargar (rclone exit=$LASTEXITCODE)"
        } else {
            Write-OK "  $sub/ descargado"
        }
    }
}

Write-Step "Listo. Verifica: Get-ChildItem $StagingDir"
