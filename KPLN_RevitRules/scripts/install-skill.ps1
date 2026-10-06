[CmdletBinding()]
param([string]$SkillDirectory, [string]$BackupRoot, [switch]$NoEnvironment)
$ErrorActionPreference = 'Stop'
$rulesRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if (-not $SkillDirectory) {
    $codexDirectory = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }
    $SkillDirectory = Join-Path $codexDirectory 'skills\kpln-revit-family-types'
}
$destination = [IO.Path]::GetFullPath($SkillDirectory)
if (Test-Path -LiteralPath $destination) {
    if (-not $BackupRoot) { $BackupRoot = Join-Path (Split-Path (Split-Path $destination -Parent) -Parent) 'skill-backups' }
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    $backup = Join-Path $BackupRoot ('kpln-revit-family-types-' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff'))
    Copy-Item -LiteralPath $destination -Destination $backup -Recurse
    Write-Output "Skill backup: $backup"
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $rulesRoot 'skills\kpln-revit-family-types\SKILL.md') -Destination (Join-Path $destination 'SKILL.md')
[IO.File]::WriteAllText((Join-Path $destination 'rules-root.txt'), $rulesRoot, [Text.UTF8Encoding]::new($false))
if (-not $NoEnvironment) { [Environment]::SetEnvironmentVariable('KPLN_REVIT_RULES_ROOT', $rulesRoot, 'User') }
Write-Output "Installed skill: $destination"
Write-Output "Rules checkout: $rulesRoot"
