[CmdletBinding()]
param(
    [ValidateSet('2020','2023','2024','2026')][string[]]$RevitVersion = @('2020','2023','2024','2026'),
    [switch]$DebugBuild,
    [string]$MSBuildPath
)
$ErrorActionPreference = 'Stop'
# Classic WPF Resource items require the .NET Framework MSBuild used by Visual Studio.
if (-not $MSBuildPath) {
    $command = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if ($command) { $MSBuildPath = $command.Source }
    else {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
        if (Test-Path -LiteralPath $vswhere) {
            $MSBuildPath = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\Current\Bin\MSBuild.exe' | Select-Object -First 1
        }
    }
}
if (($RevitVersion | Where-Object { $_ -ne '2026' }) -and (-not $MSBuildPath -or -not (Test-Path -LiteralPath $MSBuildPath))) {
    throw 'Visual Studio / Build Tools MSBuild.exe is required for WPF image resources. Install the .NET desktop development workload or pass -MSBuildPath.'
}
$project = Join-Path $PSScriptRoot '..\KPLN_RevitMcpBridge.csproj'
foreach ($version in $RevitVersion) {
    $configuration = if ($DebugBuild) { 'Debug' + $version } else { 'Revit' + $version }
    if ($version -eq '2026') {
        & dotnet build $project -c $configuration /p:Platform=x64 /v:minimal /nologo
        if ($LASTEXITCODE -ne 0) { throw "Build failed: $configuration" }
        continue
    }
    & $MSBuildPath $project "/p:Configuration=$configuration" /p:Platform=x64 /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $configuration" }
}
