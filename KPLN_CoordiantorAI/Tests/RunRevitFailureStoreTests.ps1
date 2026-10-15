param(
    [string]$CompilerPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe",
    [string]$RuntimeJsonAssembly
)
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $projectRoot "obj\RevitFailureStoreTests"
$jsonAssembly = Join-Path $projectRoot "packages\Newtonsoft.Json.13.0.4\lib\net45\Newtonsoft.Json.dll"
if ($RuntimeJsonAssembly) {
    $RuntimeJsonAssembly = (Resolve-Path -LiteralPath $RuntimeJsonAssembly).Path
    $outputDirectory += "_RuntimeJson"
}
else {
    $RuntimeJsonAssembly = $jsonAssembly
}
$testExe = Join-Path $outputDirectory "RevitFailureStoreTests.exe"
$sources = @(
    "ExternalAIModel\Mcp\RevitFailureStore.cs",
    "ExternalAIModel\Mcp\RevitMcpContracts.cs",
    "ExternalAIModel\Mcp\RevitMcpToolRegistry.cs",
    "ExternalAIModel\Mcp\RevitMcpTransport.cs",
    "ExternalAIModel\Mcp\RevitMcpSecurity.cs",
    "Tests\RevitFailureStoreTests.cs"
) | ForEach-Object { Join-Path $projectRoot $_ }
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
& $CompilerPath /nologo /langversion:7.3 /target:exe /platform:x64 /r:System.Net.Http.dll "/r:$jsonAssembly" "/out:$testExe" $sources
if ($LASTEXITCODE -ne 0) { throw "Test compilation failed." }
Copy-Item -LiteralPath $RuntimeJsonAssembly -Destination (Join-Path $outputDirectory "Newtonsoft.Json.dll") -Force
Write-Output ("Runtime Newtonsoft.Json: " + (Get-Item -LiteralPath $RuntimeJsonAssembly).VersionInfo.FileVersion)
& $testExe
if ($LASTEXITCODE -ne 0) { throw "Failure capture tests failed." }
