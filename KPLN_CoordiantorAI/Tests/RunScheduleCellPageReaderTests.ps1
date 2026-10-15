param(
    [string]$CompilerPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $projectRoot "obj\ScheduleCellPageReaderTests"
$jsonAssembly = Join-Path $projectRoot "packages\Newtonsoft.Json.13.0.4\lib\net45\Newtonsoft.Json.dll"
$reader = Join-Path $projectRoot "ExternalAIModel\ScheduleCellPageReader.cs"
$testSource = Join-Path $PSScriptRoot "ScheduleCellPageReaderTests.cs"
$testExe = Join-Path $outputDirectory "ScheduleCellPageReaderTests.exe"

if (!(Test-Path -LiteralPath $CompilerPath) -or !(Test-Path -LiteralPath $jsonAssembly)) {
    throw "The C# compiler or restored Newtonsoft.Json package is missing."
}
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
& $CompilerPath /nologo /langversion:7.3 /target:exe /platform:x64 "/r:$jsonAssembly" "/out:$testExe" $reader $testSource
if ($LASTEXITCODE -ne 0) { throw "Test compilation failed." }
Copy-Item -LiteralPath $jsonAssembly -Destination $outputDirectory -Force
& $testExe
if ($LASTEXITCODE -ne 0) { throw "Schedule cell pagination tests failed." }
