[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Python)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
& $Python -B -m unittest discover -s (Join-Path $root 'tests') -v
if ($LASTEXITCODE -ne 0) { throw 'MCP tests failed' }
foreach ($suite in @('Dispatcher','Transport','Parameters','FamilyTypes','Lifecycle','Inspection','Publication')) {
    & dotnet msbuild (Join-Path $root "tests\$suite\$suite.csproj") /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $suite" }
    $exe = switch ($suite) { 'Dispatcher' { 'DispatcherTests.exe' }; 'Transport' { 'TransportTests.exe' }; 'Parameters' { 'ParameterTests.exe' }; 'FamilyTypes' { 'FamilyTypeTests.exe' }; 'Lifecycle' { 'LifecycleTests.exe' }; 'Inspection' { 'InspectionTests.exe' } }
    if ($suite -eq 'Publication') { $exe = 'KPLN_Publication.exe' }
    & (Join-Path $root "tests\$suite\bin\$exe")
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $suite" }

    & dotnet build (Join-Path $root 'tests\Revit2026\Revit2026.csproj') "/p:Suite=$suite" /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Revit 2026 test build failed: $suite" }
    $assembly = if ($suite -eq 'Publication') { 'KPLN_Publication' } else { 'Revit2026' }
    & dotnet (Join-Path $root "tests\Revit2026\bin\$suite\net8.0-windows\$assembly.dll")
    if ($LASTEXITCODE -ne 0) { throw "Revit 2026 test failed: $suite" }
}
