[CmdletBinding()]
param(
    [string]$RulesRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$SkillDirectory,
    [switch]$NoEnvironment
)

$ErrorActionPreference = 'Stop'
$rulesPath = (Resolve-Path -LiteralPath $RulesRoot).Path
foreach ($relative in @('VERSION', 'rules\clash-policy.md', 'src\classify_opening_clashes.py')) {
    if (-not (Test-Path -LiteralPath (Join-Path $rulesPath $relative) -PathType Leaf)) {
        throw "Incomplete rules checkout: $relative"
    }
}
$sourceSkill = Join-Path $rulesPath 'skills\navisworks-clash-tolerances'
if ([string]::IsNullOrWhiteSpace($SkillDirectory)) {
    $skillHome = $env:CODEX_HOME
    if ([string]::IsNullOrWhiteSpace($skillHome)) {
        $skillHome = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.codex'
    }
    $SkillDirectory = Join-Path $skillHome 'skills\navisworks-clash-tolerances'
}
$destination = [IO.Path]::GetFullPath($SkillDirectory)
if ($destination.TrimEnd('\') -eq $sourceSkill.TrimEnd('\')) {
    throw 'The installed skill must be separate from its repository source.'
}
$backupPath = $null
if (Test-Path -LiteralPath $destination) {
    $backupPath = Join-Path (Split-Path -Parent (Split-Path -Parent $destination)) (
        'skill-backups\navisworks-clash-tolerances-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss-fffffff'))
    New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
    Copy-Item -LiteralPath $destination -Destination (Join-Path $backupPath 'skill') -Recurse
}
New-Item -ItemType Directory -Path $destination -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $sourceSkill 'SKILL.md') -Destination (Join-Path $destination 'SKILL.md')
$rulesPath | Set-Content -LiteralPath (Join-Path $destination 'rules-root.txt') -Encoding UTF8
if (-not (Test-Path -LiteralPath (Join-Path $destination 'agents\openai.yaml'))) {
    New-Item -ItemType Directory -Path (Join-Path $destination 'agents') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $sourceSkill 'agents\openai.yaml') -Destination (Join-Path $destination 'agents\openai.yaml')
}
# The former sources are backed up above; active policy now lives in the checkout.
foreach ($relative in @('references\clash-policy.md', 'scripts\classify_opening_clashes.py')) {
    $legacyFile = Join-Path $destination $relative
    if (Test-Path -LiteralPath $legacyFile -PathType Leaf) {
        Remove-Item -LiteralPath $legacyFile
    }
}
if (-not $NoEnvironment) {
    [Environment]::SetEnvironmentVariable('KPLN_NAVIS_RULES_ROOT', $rulesPath, 'User')
    $env:KPLN_NAVIS_RULES_ROOT = $rulesPath
}
[pscustomobject]@{
    RulesRoot = $rulesPath
    InstalledSkill = $destination
    Backup = $backupPath
    UserEnvironmentUpdated = (-not $NoEnvironment)
}
