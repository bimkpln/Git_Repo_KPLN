param([string]$Launcher, [string]$Workspace)
$ErrorActionPreference = 'Stop'
# Exercise the normal branch without opening a real app/chat.
function Start-Process {
    param([string]$FilePath)
    Write-Output ('MOCK_URI:' + $FilePath)
}
& $Launcher -Workspace $Workspace
