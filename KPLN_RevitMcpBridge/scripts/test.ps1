[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Python)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
& $Python -B -m unittest discover -s (Join-Path $root 'tests') -v
if ($LASTEXITCODE -ne 0) { throw 'MCP tests failed' }
foreach ($suite in @('Dispatcher','Transport','Parameters','FamilyTypes','Lifecycle')) {
    & dotnet msbuild (Join-Path $root "tests\$suite\$suite.csproj") /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $suite" }
    $exe = switch ($suite) { 'Dispatcher' { 'DispatcherTests.exe' }; 'Transport' { 'TransportTests.exe' }; 'Parameters' { 'ParameterTests.exe' }; 'FamilyTypes' { 'FamilyTypeTests.exe' }; 'Lifecycle' { 'LifecycleTests.exe' } }
    & (Join-Path $root "tests\$suite\bin\$exe")
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $suite" }
}
