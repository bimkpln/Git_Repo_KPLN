[CmdletBinding()]
param([string]$SkillDirectory, [switch]$NoEnvironment)
$ErrorActionPreference = 'Stop'
$rulesRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if (-not $SkillDirectory) {
    $codexDirectory = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }
    $SkillDirectory = Join-Path $codexDirectory 'skills\kpln-revit-bridge'
}
$destination = [IO.Path]::GetFullPath($SkillDirectory)
if (Test-Path -LiteralPath $destination) {
    $backup = $destination + '.backup-' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff')
    Copy-Item -LiteralPath $destination -Destination $backup -Recurse
    Write-Output "Skill backup: $backup"
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $rulesRoot 'skills\kpln-revit-bridge\SKILL.md') -Destination (Join-Path $destination 'SKILL.md')
[IO.File]::WriteAllText((Join-Path $destination 'rules-root.txt'), $rulesRoot, [Text.UTF8Encoding]::new($false))
if (-not $NoEnvironment) { [Environment]::SetEnvironmentVariable('KPLN_REVIT_RULES_ROOT', $rulesRoot, 'User') }
Write-Output "Installed skill: $destination"
Write-Output "Rules checkout: $rulesRoot"
