param(
    [string]$CompilerPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $projectRoot "obj\ModelChatMarkdownRendererTests"
$frameworkRoot = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
$renderer = Join-Path $projectRoot "Forms\ModelChatMarkdownRenderer.cs"
$testSource = Join-Path $PSScriptRoot "ModelChatMarkdownRendererTests.cs"
$testExe = Join-Path $outputDirectory "ModelChatMarkdownRendererTests.exe"

if (!(Test-Path -LiteralPath $CompilerPath) -or !(Test-Path -LiteralPath $frameworkRoot)) {
    throw "The C# compiler or .NET Framework 4.8 reference assemblies are missing."
}

New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
& $CompilerPath /nologo /langversion:7.3 /target:exe /platform:x64 "/r:$(Join-Path $frameworkRoot 'WindowsBase.dll')" "/r:$(Join-Path $frameworkRoot 'PresentationCore.dll')" "/r:$(Join-Path $frameworkRoot 'PresentationFramework.dll')" "/r:$(Join-Path $frameworkRoot 'System.Xaml.dll')" "/out:$testExe" $renderer $testSource
if ($LASTEXITCODE -ne 0) { throw "Markdown renderer test compilation failed." }

& $testExe
if ($LASTEXITCODE -ne 0) { throw "Markdown renderer tests failed." }
