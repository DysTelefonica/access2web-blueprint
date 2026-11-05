# Collects what classify-tables.py needs to classify every table of the phase 0 reference copy
# (docs/07-migracion/fase-0/clasificacion-tablas.md). For each application in classification-bases.tsv
# it extracts src/ from its base ref and reads its front binary from that ref (never the working tree),
# then reads the saved queries of every reference backend and Seguridad's data-driven table catalog.
# Every database is opened through DAO in read-only mode on a temporary copy.
param(
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$Repos = 'C:\00repos\codigo',
    [string]$Reference = 'C:\00repos\access2web-datos',
    [string]$Bases
)
$ErrorActionPreference = 'Stop'
if (-not $env:ACCDB_PASSWORD) { throw 'ACCDB_PASSWORD is not set' }
if (-not $Bases) { $Bases = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'classification-bases.tsv' }
New-Item -ItemType Directory -Force $OutDir | Out-Null
Copy-Item -LiteralPath $Bases (Join-Path $OutDir 'bases.tsv')
$engine = New-Object -ComObject DAO.DBEngine.120
$frontRows = New-Object System.Collections.Generic.List[string]
$frontRows.Add("app`tname`tkind`ttarget_backend`tsource_table")
$queryRows = New-Object System.Collections.Generic.List[string]
$queryRows.Add("owner`tquery`tsql")

function Open-Copy([string]$path) {
    $work = Join-Path ([IO.Path]::GetTempPath()) ('cu-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory $work | Out-Null
    $copy = Join-Path $work 'db.accdb'
    Copy-Item -LiteralPath $path $copy
    try { $db = $engine.OpenDatabase($copy, $false, $true, ';PWD=' + $env:ACCDB_PASSWORD) }
    catch { $db = $engine.OpenDatabase($copy, $false, $true, '') }
    , @($db, $work)
}

function Add-Queries($db, [string]$owner) {
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -like '~*') { continue }
        $queryRows.Add("$owner`t$($q.Name)`t$(([string]$q.SQL) -replace "[`r`n`t]+", ' ')")
    }
}

foreach ($line in Get-Content -LiteralPath $Bases -Encoding UTF8) {
    if (-not $line.Trim()) { continue }
    $app, $repo, $ref = $line.Split("`t")
    $repoPath = Join-Path $Repos $repo
    git -C $repoPath fetch -q origin
    $srcOut = Join-Path $OutDir "src\$repo"
    New-Item -ItemType Directory -Force $srcOut | Out-Null
    $tar = Join-Path $OutDir "$repo.tar"
    git -C $repoPath archive -o $tar $ref src
    # Windows tar: Git's GNU tar reads 'C:' as a remote host.
    & (Join-Path $env:SystemRoot 'System32\tar.exe') -x -f $tar -C $srcOut
    Remove-Item -LiteralPath $tar

    $binary = git -C $repoPath ls-tree --name-only $ref | Where-Object { $_ -match '\.accdb$' -and $_ -notmatch '(?i)datos' } | Select-Object -First 1
    $frontCopy = Join-Path ([IO.Path]::GetTempPath()) ('front-' + [Guid]::NewGuid().ToString('N') + '.accdb')
    if ($binary) {
        Start-Process -FilePath git -ArgumentList @('-C', $repoPath, 'show', "${ref}:$binary") -RedirectStandardOutput $frontCopy -NoNewWindow -Wait | Out-Null
    } else {
        # The front is not versioned on its base ref (AGEDYS ECO): read the checkout's binary instead.
        $local = Get-ChildItem -LiteralPath $repoPath -Filter *.accdb -File | Where-Object { $_.Name -notmatch '(?i)datos' } | Select-Object -First 1
        if (-not $local) { $frontRows.Add("$app`t-`terror`t`tno front binary"); continue }
        Copy-Item -LiteralPath $local.FullName $frontCopy
    }
    $opened = Open-Copy $frontCopy
    $db = $opened[0]
    try {
        foreach ($tableDef in $db.TableDefs) {
            if ($tableDef.Name -like 'MSys*' -or $tableDef.Name -like '~*' -or $tableDef.Name -like 'USys*') { continue }
            if ($tableDef.Connect -match 'DATABASE=([^;]+)') { $frontRows.Add("$app`t$($tableDef.Name)`tlink`t$(Split-Path $Matches[1] -Leaf)`t$($tableDef.SourceTableName)") }
            elseif (-not $tableDef.Connect) { $frontRows.Add("$app`t$($tableDef.Name)`tlocal`t`t") }
        }
        Add-Queries $db "front:$app"
    } finally {
        $db.Close()
        Remove-Item -Recurse -Force $opened[1]
        Remove-Item -Force $frontCopy
    }
}

Get-ChildItem -LiteralPath $Reference -Filter *.accdb -File | Where-Object { $_.Name -notmatch '\.(bak|BACKUP)-' } | ForEach-Object {
    $opened = Open-Copy $_.FullName
    $db = $opened[0]
    try {
        Add-Queries $db ('backend:' + $_.BaseName)
        if ($_.Name -eq 'Seguridad_datos.accdb') {
            $rs = $db.OpenRecordset('SELECT NombreTabla FROM Tb0TipoRegistros', 4)
            while (-not $rs.EOF) { $queryRows.Add("data:Seguridad`tTb0TipoRegistros.NombreTabla`tSELECT * FROM [$($rs.Fields(0).Value)]"); $rs.MoveNext() }
            $rs.Close()
        }
    } finally {
        $db.Close()
        Remove-Item -Recurse -Force $opened[1]
    }
}
$utf8 = New-Object System.Text.UTF8Encoding $false
[IO.File]::WriteAllLines((Join-Path $OutDir 'fronts.tsv'), $frontRows, $utf8)
[IO.File]::WriteAllLines((Join-Path $OutDir 'queries.tsv'), $queryRows, $utf8)
"fronts: $($frontRows.Count - 1) rows; queries: $($queryRows.Count - 1) rows written to $OutDir"
