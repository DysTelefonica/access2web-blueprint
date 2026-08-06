# Script: setup-staging.ps1
# Propósito: copia self-contained los 8 stagings de las apps legacy al repo access2web-blueprint/data/staging/
# Uso: pwsh -File scripts/setup-staging.ps1
#
# Convenciones:
# - Cada app en data/staging/<app>/{frontend,backend,src,docs,resources}
# - Source-of-truth: staging si existe, si no main
# - Hace skips si ya existe (idempotente; --force re-copia)
# - Los .accdb pueden contener secretos (D92, D98, D104, D109). El usuario autoriza commitearlos al repo privado.

param(
    [switch]$Force = $false,
    [string]$RepoRoot = "C:\00repos\codigo\access2web-blueprint"
)

$ErrorActionPreference = "Stop"
$StagingDir = Join-Path $RepoRoot "data\staging"

# Mapa (app, source_dir, frontend_file, backend_file)
# source_dir: staging > main (source of truth per methodology)
$Apps = @(
    @{ Id = "condor";              Source = "C:\00repos\codigo\00_CONDOR\staging";                  Frontend = "CONDOR.accdb";       Backend = "condor_datos.accdb" },
    @{ Id = "hps";                 Source = "C:\00repos\codigo\00_HPS\staging";                     Frontend = "HPS.accdb";          Backend = "HPST.accdb" },
    @{ Id = "hps-solicitudes";     Source = "C:\00repos\codigo\HPS_SOLICITUDES";                   Frontend = "Solicitudes_HPS.accdb"; Backend = "Solicitudes_HPS_datos.accdb" },
    @{ Id = "brass";               Source = "C:\00repos\codigo\00_BRASS\00_main";                   Frontend = "";                   Backend = "Gestion_Brass_Gestion.accdb" },
    @{ Id = "gestion-riesgos";     Source = "C:\00repos\codigo\00_GESTION_RIESGOS\staging";         Frontend = "Gestion_Riesgos.accdb"; Backend = "Gestion_Riesgos_Datos.accdb" },
    @{ Id = "no-conformidades";    Source = "C:\00repos\codigo\00_NO_CONFORMIDADES\staging";        Frontend = "NoConformidades.accdb"; Backend = "NoConformidades_Datos.accdb" },
    @{ Id = "expedientes";         Source = "C:\00repos\codigo\00_EXPEDIENTES\staging";            Frontend = "Expedientes.accdb";    Backend = "Expedientes_datos.accdb" },
    @{ Id = "lanzaderas";          Source = "C:\00repos\codigo\00_LANZADERA\staging";               Frontend = "Lanzadera.accdb";     Backend = "Lanzadera_Datos.accdb" }
)

function Write-Step($msg) { Write-Host "[SETUP-STAGING] $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "[SKIP] $msg" -ForegroundColor Yellow }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

foreach ($app in $Apps) {
    $appId = $app.Id
    $source = $app.Source
    $appDir = Join-Path $StagingDir $appId

    Write-Step "Procesando $appId (source: $source)"

    if (-not (Test-Path $source)) {
        Write-Warn "  Source no existe: $source — SKIP"
        continue
    }

    foreach ($sub in @("frontend", "backend", "src", "docs", "resources")) {
        $dest = Join-Path $appDir $sub
        if ((Test-Path $dest) -and -not $Force) {
            Write-Skip "  $sub ya existe — SKIP (usa -Force para re-copiar)"
            continue
        }
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
    }

    # Frontend
    if ($app.Frontend) {
        $srcPath = Join-Path $source $app.Frontend
        if (Test-Path $srcPath) {
            $destPath = Join-Path $appDir "frontend\$($app.Frontend)"
            if ((Test-Path $destPath) -and -not $Force) {
                Write-Skip "  frontend ya copiado"
            } else {
                Copy-Item $srcPath $destPath -Force
                Write-OK "  frontend copiado: $($app.Frontend)"
            }
        } else {
            Write-Warn "  frontend no encontrado: $srcPath"
        }
    }

    # Backend
    if ($app.Backend) {
        $srcPath = Join-Path $source $app.Backend
        if (Test-Path $srcPath) {
            $destPath = Join-Path $appDir "backend\$($app.Backend)"
            if ((Test-Path $destPath) -and -not $Force) {
                Write-Skip "  backend ya copiado"
            } else {
                Copy-Item $srcPath $destPath -Force
                Write-OK "  backend copiado: $($app.Backend)"
            }
        } else {
            Write-Warn "  backend no encontrado: $srcPath"
        }
    }

    # src/
    $srcSrc = Join-Path $source "src"
    if (Test-Path $srcSrc) {
        $destSrc = Join-Path $appDir "src"
        if ((Test-Path $destSrc) -and (Get-ChildItem $destSrc | Measure-Object).Count -gt 0 -and -not $Force) {
            Write-Skip "  src/ ya copiado"
        } else {
            Copy-Item -Path "$srcSrc\*" -Destination $destSrc -Recurse -Force
            Write-OK "  src/ copiado"
        }
    } else {
        Write-Warn "  src/ no encontrado en $source"
    }

    # docs/ (opcional; puede no existir)
    $srcDocs = Join-Path $source "docs"
    if (Test-Path $srcDocs) {
        $destDocs = Join-Path $appDir "docs"
        if ((Test-Path $destDocs) -and (Get-ChildItem $destDocs -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0 -and -not $Force) {
            Write-Skip "  docs/ ya copiado"
        } else {
            Copy-Item -Path "$srcDocs\*" -Destination $destDocs -Recurse -Force
            Write-OK "  docs/ copiado"
        }
    }

    # resources/ (opcional)
    $srcRes = Join-Path $source "resources"
    if (Test-Path $srcRes) {
        $destRes = Join-Path $appDir "resources"
        if ((Test-Path $destRes) -and (Get-ChildItem $destRes -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0 -and -not $Force) {
            Write-Skip "  resources/ ya copiado"
        } else {
            Copy-Item -Path "$srcRes\*" -Destination $destRes -Recurse -Force
            Write-OK "  resources/ copiado"
        }
    }

    # clases/, forms/, modules/, queries/, reports/ (algunos repos los tienen al mismo nivel que src/)
    foreach ($sub in @("classes", "forms", "modules", "queries", "reports")) {
        $srcSub = Join-Path $source $sub
        if (Test-Path $srcSub) {
            $destSrc = Join-Path $appDir "src"
            New-Item -ItemType Directory -Path $destSrc -Force | Out-Null
            $destSub = Join-Path $destSrc $sub
            if ((Test-Path $destSub) -and -not $Force) {
                Write-Skip "  src/$sub/ ya copiado"
            } else {
                Copy-Item -Path "$srcSub\*" -Destination $destSub -Recurse -Force
                Write-OK "  src/$sub/ copiado"
            }
        }
    }
}

Write-Step "Listo. Verifica: Get-ChildItem $StagingDir"
