# Data inventory of the phase 0 reference copy (docs/07-migracion/fase-0/copia-referencia.md).
# Verifies every backend against the SHA-256 manifest, then works on temporary copies opened
# with DAO in read-only mode; the reference files are never opened. Output is tab-separated:
# one row per local table, plus one row per key candidate of tables without a primary key.
# With -Mode Ownership it lists every table and link instead: backend, name, local or link,
# and for a link the backend and table it points to (docs/07-migracion/fase-0/mapa-propiedad.md).
# With -Mode Columns it lists every column of every local table: DAO type, size, attributes,
# whether it is a complex (multivalue or attachment) field, required, validation rule and default
# value (docs/07-migracion/fase-0/coste-cambios.md).
param(
    [string]$Root = 'C:\00repos\access2web-datos',
    [string]$Manifest,
    [ValidateSet('Data', 'Ownership', 'Columns')][string]$Mode = 'Data',
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
if (-not $Manifest) { $Manifest = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) '..\..\docs\07-migracion\fase-0\copia-referencia.md' }
if (-not $env:ACCDB_PASSWORD) { throw 'ACCDB_PASSWORD is not set' }

$expected = @{}
foreach ($line in Get-Content -LiteralPath $Manifest -Encoding UTF8) {
    if ($line -match '^\| `([^`]+\.accdb)` \| \d+ \| `([0-9a-f]{64})` \|$') { $expected[$Matches[1]] = $Matches[2] }
}
if ($expected.Count -eq 0) { throw "no manifest rows in $Manifest" }
$sha = [Security.Cryptography.SHA256]::Create()
foreach ($name in $expected.Keys) {
    $stream = [IO.File]::OpenRead((Join-Path $Root $name))
    try { $actual = -join ($sha.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) } finally { $stream.Dispose() }
    if ($actual -ne $expected[$name]) { throw "hash mismatch: $name" }
}

$engine = New-Object -ComObject DAO.DBEngine.120
$rows = New-Object System.Collections.Generic.List[string]
if ($Mode -eq 'Ownership') { $rows.Add("backend`tname`tkind`ttarget_backend`tsource_table") }
elseif ($Mode -eq 'Columns') { $rows.Add("backend`ttable`tcolumn`ttype`tsize`tattributes`tcomplex`trequired`tvalidation`tdefault") }
else { $rows.Add("kind`tbackend`ttable`trows`tpk`tcolumn`tnulls`tdistinct") }
foreach ($name in ($expected.Keys | Sort-Object)) {
    $work = Join-Path ([IO.Path]::GetTempPath()) ('inv-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory $work | Out-Null
    $copy = Join-Path $work $name
    Copy-Item -LiteralPath (Join-Path $Root $name) -Destination $copy
    $db = $engine.OpenDatabase($copy, $false, $true, ';PWD=' + $env:ACCDB_PASSWORD)
    try {
        foreach ($t in $db.TableDefs) {
            if ($t.Name -like 'MSys*' -or $t.Name -like '~*') { continue }
            if ($Mode -eq 'Ownership') {
                if ($t.Connect -match 'DATABASE=([^;]+)') { $rows.Add("$name`t$($t.Name)`tlink`t$(Split-Path $Matches[1] -Leaf)`t$($t.SourceTableName)") }
                else { $rows.Add("$name`t$($t.Name)`tlocal`t`t") }
                continue
            }
            if ($t.Connect) { continue }
            if ($Mode -eq 'Columns') {
                foreach ($fld in $t.Fields) {
                    $complex = $false
                    try { $complex = [bool]$fld.IsComplex } catch { $complex = $false }
                    $rule = ($fld.ValidationRule -replace '[\t\r\n]', ' ')
                    $default = ($fld.DefaultValue -replace '[\t\r\n]', ' ')
                    $rows.Add("$name`t$($t.Name)`t$($fld.Name)`t$($fld.Type)`t$($fld.Size)`t$($fld.Attributes)`t$complex`t$($fld.Required)`t$rule`t$default")
                }
                continue
            }
            $q = '[' + $t.Name + ']'
            $count = $db.OpenRecordset("SELECT COUNT(*) FROM $q", 4).Fields(0).Value
            $pk = ''
            foreach ($i in $t.Indexes) { if ($i.Primary) { $pk = (@($i.Fields) | ForEach-Object { $_.Name }) -join '+' } }
            $rows.Add("table`t$name`t$($t.Name)`t$count`t$pk`t`t`t")
            if ($pk) { continue }
            # Key candidates: autonumber fields first, then fields named ID*.
            $candidates = @($t.Fields | Where-Object { ($_.Attributes -band 16) -ne 0 }) +
                          @($t.Fields | Where-Object { ($_.Attributes -band 16) -eq 0 -and $_.Name -like 'ID*' })
            foreach ($f in ($candidates | Select-Object -First 3)) {
                $c = '[' + $f.Name + ']'
                $nulls = $db.OpenRecordset("SELECT COUNT(*) FROM $q WHERE $c IS NULL", 4).Fields(0).Value
                $distinct = $db.OpenRecordset("SELECT COUNT(*) FROM (SELECT DISTINCT $c FROM $q WHERE $c IS NOT NULL)", 4).Fields(0).Value
                $rows.Add("candidate`t$name`t$($t.Name)`t$count`t`t$($f.Name)`t$nulls`t$distinct")
            }
        }
    } finally {
        $db.Close()
        Remove-Item -Recurse -Force $work
    }
}
[IO.File]::WriteAllLines($Out, $rows, (New-Object System.Text.UTF8Encoding $false))
"$($rows.Count - 1) rows written to $Out"
