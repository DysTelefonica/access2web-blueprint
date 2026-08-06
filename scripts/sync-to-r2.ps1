# Script: sync-to-r2.ps1
# Propósito: sincroniza los binarios legacy (.accdb) y los índices codegraph-vba desde
# data/staging/ del repo a Cloudflare R2 (bucket access2web-staging-binaries).
#
# Úsalo cuando:
#   - Cambias binarios en data/staging/<app>/{frontend,backend,resources,.codegraph-vba}/
#   - Re-generas el codegraph-vba/ index
#   - Actualizas desde el staging del repo original
#
# Pre-requisitos:
#   - rclone configurado con el remote "cloudflare-r2" apuntando al bucket correcto.
#   - El bucket "access2web-staging-binaries" creado en R2.
#
# Uso:
#   pwsh -File scripts/sync-to-r2.ps1                     # todas las 8 apps
#   pwsh -File scripts/sync-to-r2.ps1 -App condor         # solo Condor
#   pwsh -File scripts/sync-to-r2.ps1 -DryRun            # ver qué subiría
#   pwsh -File scripts/sync-to-r2.ps1 -NoVerify          # skip verificación post-upload
#   pwsh -File scripts/sync-to-r2.ps1 -VerboseRclone      # ver el detalle de rclone
#
# Output:
#   - Progress bar de rclone
#   - Resumen final con bytes transferidos, errores, tiempo
#   - Exit code 0 si OK, 1 si algún error

param(
    [string]$App = "",                 # '' = todas, 'condor'/'hps'/etc = una sola
    [switch]$DryRun = $false,
    [switch]$NoVerify = $false,
    [switch]$VerboseRclone = $false,
    [string]$RepoRoot = "C:\00repos\codigo\access2web-blueprint"
)

$ErrorActionPreference = "Stop"
$StagingDir = Join-Path $RepoRoot "data\staging"
$Remote = "cloudflare-r2:access2web-staging-binaries"
$Apps = @("condor", "hps", "hps-solicitudes", "brass", "gestion-riesgos", "no-conformidades", "expedientes", "lanzaderas")

# Crear logs/ si no existe
$logsDir = Join-Path $RepoRoot "logs"
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }

function Write-Step($msg) { Write-Host "[SYNC-R2] $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "[SKIP] $msg" -ForegroundColor Yellow }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

# 1. Validar argumentos
if ($App -ne "" -and $Apps -notcontains $App) {
    Write-Err "App '$App' no reconocida. Apps válidas: $($Apps -join ', ')"
    exit 1
}

# 2. Verificar rclone + remote
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

# 3. Verificar acceso al bucket
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

# 4. Determinar targets
$targets = if ($App -ne "") { @($App) } else { $Apps }

# 5. Stats
$totalErrors = 0
$sw = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($appId in $targets) {
    $appDir = Join-Path $StagingDir $appId
    if (-not (Test-Path $appDir)) {
        Write-Warn "  $appDir no existe localmente. SKIP."
        continue
    }
    Write-Step "Subiendo $appId..."

    foreach ($sub in @("frontend", "backend", "resources", ".codegraph-vba")) {
        $src = Join-Path $appDir $sub
        if (-not (Test-Path $src)) {
            Write-Skip "  $sub/ no existe localmente. SKIP."
            continue
        }

        $rcloneArgs = @(
            "copy",
            "$src/",
            "$Remote/$appId/$sub/",
            "--transfers", "4",
            "--checkers", "8",
            "--stats", "30s"
        )
        if ($DryRun) {
            $rcloneArgs += "--dry-run"
        }
        if ($VerboseRclone) {
            $rcloneArgs += "--verbose"
        } else {
            $rcloneArgs += "--progress"
        }

        $logFile = Join-Path $logsDir "rclone-sync-$appId-$sub.log"
        $errFile = Join-Path $logsDir "rclone-sync-$appId-$sub.err"
        Write-Host "  -> rclone copy $src/ $Remote/$appId/$sub/" -ForegroundColor Gray
        $proc = Start-Process -FilePath "rclone" -ArgumentList $rcloneArgs -NoNewWindow -PassThru -Wait `
            -RedirectStandardOutput $logFile `
            -RedirectStandardError $errFile
        $exitCode = $proc.ExitCode

        if ($exitCode -ne 0) {
            Write-Err "  $sub/ falló (rclone exit=$exitCode). Ver $errFile"
            $totalErrors++
            continue
        }

        # Extraer "Transferred:" del log para feedback
        if (Test-Path $logFile) {
            $logContent = Get-Content $logFile -Raw
            $transferredLine = ($logContent | Select-String -Pattern "Transferred:").ToString().Trim()
            if ($transferredLine) {
                Write-OK "  $sub/ $transferredLine"
            } else {
                Write-OK "  $sub/ subido"
            }
        } else {
            Write-OK "  $sub/ subido"
        }
    }
}

$sw.Stop()
$totalSec = [int]$sw.Elapsed.TotalSeconds
$minutes = [math]::Floor($totalSec / 60)
$seconds = $totalSec % 60

# 6. Verificación post-upload (opcional)
if (-not $NoVerify -and -not $DryRun) {
    Write-Step "Verificación post-upload..."
    foreach ($appId in $targets) {
        $appDir = Join-Path $StagingDir $appId
        if (-not (Test-Path $appDir)) { continue }
        foreach ($sub in @("frontend", "backend", "resources", ".codegraph-vba")) {
            $src = Join-Path $appDir $sub
            if (-not (Test-Path $src)) { continue }
            & rclone check "$src/" "$Remote/$appId/$sub/" --quiet 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warn "  $appId/$sub/ verificación irregular (puede ser normal si hay cambios en curso)"
            }
        }
    }
}

# 7. Resumen final
Write-Host ""
Write-Host "========== RESUMEN ==========" -ForegroundColor Magenta
Write-Host "Apps procesadas: $($targets -join ', ')"
Write-Host "Tiempo: ${minutes}m ${seconds}s"
Write-Host "Errores: $totalErrors"
Write-Host "=============================" -ForegroundColor Magenta

if ($totalErrors -gt 0) {
    exit 1
}
exit 0
