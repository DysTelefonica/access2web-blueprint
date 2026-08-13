# Installs the repository skills into the agent directories present locally.
#
# Windows counterpart of install-skills.sh. It copies rather than symlinks:
# symlinks on Windows need developer mode or elevation, and a copy behaves the
# same on every machine.
#
# Re-run after any `git pull` that touches skills/.
#
# Contract: skills/README.md
#
#     powershell -ExecutionPolicy Bypass -File scripts/install-skills.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $repoRoot 'skills'

if (-not (Test-Path -LiteralPath $sourceDir)) {
    Write-Error "install-skills: $sourceDir does not exist"
}

# A target is written only when its parent already exists, so nobody gets a
# directory for an agent they do not use.
$targets = [ordered]@{
    '.claude'  = 'Claude Code'
    '.opencode' = 'OpenCode'
    '.codex'   = 'Codex'
}

# Counted apart on purpose: "no agent installed here" and "no skills to install
# yet" are different situations, and reporting one as the other sends the reader
# to fix something that is not broken.
$agentsFound = 0
$skillsCopied = 0

foreach ($dir in $targets.Keys) {
    $parent = Join-Path $repoRoot $dir
    if (-not (Test-Path -LiteralPath $parent)) { continue }
    $agentsFound++

    $dest = Join-Path $parent 'skills'
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    foreach ($skill in Get-ChildItem -LiteralPath $sourceDir -Directory) {
        $skillDest = Join-Path $dest $skill.Name
        if (Test-Path -LiteralPath $skillDest) {
            Remove-Item -LiteralPath $skillDest -Recurse -Force
        }
        Copy-Item -LiteralPath $skill.FullName -Destination $skillDest -Recurse
        Write-Host "install-skills: $($skill.Name) -> $dir/skills/$($skill.Name)"
        $skillsCopied++
    }

    Write-Host "install-skills: $($targets[$dir]) ready at $dest"
}

if ($agentsFound -eq 0) {
    Write-Host 'install-skills: no agent directory found (.claude, .opencode, .codex).'
    Write-Host 'install-skills: create the one your agent uses and run this again.'
}
elseif ($skillsCopied -eq 0) {
    Write-Host 'install-skills: skills/ holds no skill yet; nothing to copy.'
}
