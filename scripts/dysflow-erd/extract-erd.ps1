# Read-only schema extraction for an Access database (#1074).
# Works on a temporary copy opened with DAO in read-only mode; the source file is never written.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Not a file: $Path" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) ("erd-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null
$copy = Join-Path $work 'db.accdb'
Copy-Item -LiteralPath $Path -Destination $copy
$connect = ''
if ($env:ACCDB_PASSWORD) { $connect = ';PWD=' + $env:ACCDB_PASSWORD }
$engine = New-Object -ComObject DAO.DBEngine.120
try {
    $db = $engine.OpenDatabase($copy, $false, $true, $connect)
    $tables = foreach ($t in $db.TableDefs) {
        if ($t.Name -like 'MSys*' -or $t.Name -like '~*') { continue }
        [pscustomobject]@{
            name      = $t.Name
            linked    = [bool]$t.Connect
            connect   = $(if ($t.Connect) { ($t.Connect -replace 'PWD=[^;]*', 'PWD=***') } else { $null })
            source    = $t.SourceTableName
            fields    = @(foreach ($f in $t.Fields) {
                    [pscustomobject]@{ name = $f.Name; type = $f.Type; size = $f.Size; required = $f.Required; allowZeroLength = $f.AllowZeroLength; defaultValue = $f.DefaultValue }
                })
            indexes   = @(foreach ($i in $t.Indexes) {
                    [pscustomobject]@{ name = $i.Name; primary = $i.Primary; unique = $i.Unique; fields = @(foreach ($f in $i.Fields) { $f.Name }) }
                })
        }
    }
    $relations = foreach ($r in $db.Relations) {
        [pscustomobject]@{
            name = $r.Name; table = $r.Table; foreignTable = $r.ForeignTable; attributes = $r.Attributes
            fields = @(foreach ($f in $r.Fields) { [pscustomobject]@{ name = $f.Name; foreignName = $f.ForeignName } })
        }
    }
    $queries = foreach ($q in $db.QueryDefs) { if ($q.Name -notlike '~*') { [pscustomobject]@{ name = $q.Name; type = $q.Type } } }
    $forms = foreach ($d in $db.Containers['Forms'].Documents) { $d.Name }
    $result = [ordered]@{ tables = @($tables); relations = @($relations); queries = @($queries); forms = @($forms) }
    $result | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $OutDir 'erd.json')
    $db.Close()
}
finally {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($engine)
    Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
}
