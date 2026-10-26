[CmdletBinding()]
param([string]$SkillDirectory)
$ErrorActionPreference = 'Stop'
$rulesRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if (-not $SkillDirectory) {
    $codexDirectory = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }
    $SkillDirectory = Join-Path $codexDirectory 'skills\kpln-model-documentation-check'
}
$destination = [IO.Path]::GetFullPath($SkillDirectory)
if (Test-Path -LiteralPath $destination) {
    $backupRoot = Join-Path $env:USERPROFILE 'Documents\KPLN\ModelDocumentationChecks\maintenance\skill-backups'
    if ([IO.Path]::GetPathRoot([IO.Path]::GetFullPath($backupRoot)) -ne 'C:\') {
        throw 'Skill backups must stay on the user C: drive; see rules/storage.md.'
    }
    $backup = Join-Path $backupRoot ([DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff'))
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    Copy-Item -LiteralPath $destination -Destination $backup -Recurse
    Write-Output "Backup: $backup"
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$skillSource = Join-Path $rulesRoot 'skills\kpln-model-documentation-check'
Copy-Item -LiteralPath (Join-Path $skillSource 'SKILL.md') -Destination (Join-Path $destination 'SKILL.md') -Force
New-Item -ItemType Directory -Path (Join-Path $destination 'agents') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $skillSource 'agents\openai.yaml') -Destination (Join-Path $destination 'agents\openai.yaml') -Force
[IO.File]::WriteAllText((Join-Path $destination 'rules-root.txt'), $rulesRoot, [Text.UTF8Encoding]::new($false))
Write-Output "Installed: $destination"
Write-Output "Rules: $rulesRoot"
