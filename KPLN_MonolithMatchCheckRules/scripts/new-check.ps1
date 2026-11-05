[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Workspace,
    [switch]$PreviewOnly
)
$ErrorActionPreference = 'Stop'
$rulesRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
if (-not [IO.Path]::IsPathRooted($Workspace)) { throw 'Workspace must be an absolute path' }
$folder = Get-Item -LiteralPath $Workspace
if (-not $folder.PSIsContainer) { throw 'Workspace must be an existing directory' }
$fullPath = $folder.FullName
if ($fullPath.StartsWith('\\')) { throw 'Use a local workspace; model and rules may be on a network share' }
$drive = Get-PSDrive -Name ([IO.Path]::GetPathRoot($fullPath).TrimEnd('\').TrimEnd(':')) -ErrorAction Stop
if ($drive.DisplayRoot -and $drive.DisplayRoot.StartsWith('\\')) { throw 'Use a local workspace, not a mapped network drive' }
$templatePath = Join-Path $rulesRoot 'examples\start-request.txt'
$promptText = [IO.File]::ReadAllText($templatePath, [Text.Encoding]::UTF8)
$uri = 'codex://new?prompt=' + [Uri]::EscapeDataString($promptText) + '&path=' + [Uri]::EscapeDataString($fullPath)
if ($PreviewOnly) { Write-Output $uri; return }
# A supported app deep link populates the composer. It does not submit a turn.
Start-Process -FilePath $uri
