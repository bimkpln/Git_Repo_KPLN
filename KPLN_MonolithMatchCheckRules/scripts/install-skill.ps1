[CmdletBinding()]
param([string]$SkillDirectory, [string]$BackupRoot)
$ErrorActionPreference = 'Stop'
$rulesRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
if (-not $SkillDirectory) {
    $SkillDirectory = Join-Path $env:USERPROFILE '.agents\skills\kpln-monolith-match-check'
    $legacyRoot = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }
    $legacy = Join-Path $legacyRoot 'skills\kpln-monolith-match-check'
    if (Test-Path -LiteralPath $legacy) { $SkillDirectory = $legacy }
}
$destination = [IO.Path]::GetFullPath($SkillDirectory)
$source = Join-Path $rulesRoot 'skills\kpln-monolith-match-check\SKILL.md'
if (-not (Test-Path -LiteralPath $source)) { throw 'Source SKILL.md is missing' }
if (Test-Path -LiteralPath $destination) {
    if (-not $BackupRoot) { $BackupRoot = Join-Path (Split-Path (Split-Path $destination -Parent) -Parent) 'skill-backups' }
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    $backup = Join-Path $BackupRoot ('kpln-monolith-match-check-' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff'))
    Copy-Item -LiteralPath $destination -Destination $backup -Recurse
    Write-Output "Skill backup: $backup"
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
Copy-Item -LiteralPath $source -Destination (Join-Path $destination 'SKILL.md')
$metadataSource = Join-Path $rulesRoot 'skills\kpln-monolith-match-check\agents\openai.yaml'
if (Test-Path -LiteralPath $metadataSource) {
    $agentsDirectory = Join-Path $destination 'agents'
    New-Item -ItemType Directory -Path $agentsDirectory -Force | Out-Null
    Copy-Item -LiteralPath $metadataSource -Destination (Join-Path $agentsDirectory 'openai.yaml')
}
[IO.File]::WriteAllText((Join-Path $destination 'rules-root.txt'), $rulesRoot, [Text.UTF8Encoding]::new($false))
Write-Output "Installed skill: $destination"
Write-Output "Shared rules: $rulesRoot"
Write-Output 'Start a new Agent chat in Cursor or Codex; if discovery has not refreshed, restart that application.'
