param(
    [string]$CompilerPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe",
    [string]$PluginAssembly = "C:\KPLN_Debug\KPLN_CoordinatorAI\Debug\2023\KPLN_CoordiantorAI.dll",
    [string]$NativeSqliteAssembly = "X:\BIM\5_Scripts\SetupFiles\Revit\1.Common\KPLN\2023\KPLN_Loader\e_sqlite3.dll"
)
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $projectRoot "obj\DepartmentPromptTests"
$sqliteAssembly = Join-Path $projectRoot "packages\System.Data.SQLite.2.0.3\lib\net471\System.Data.SQLite.dll"
$testExe = Join-Path $outputDirectory "DepartmentPromptTests.exe"
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
& $CompilerPath /nologo /langversion:7.3 /target:exe /platform:x64 /r:System.Data.dll "/r:$PluginAssembly" "/r:$sqliteAssembly" "/out:$testExe" (Join-Path $PSScriptRoot "DepartmentPromptTests.cs")
if ($LASTEXITCODE -ne 0) { throw "Test compilation failed." }
Copy-Item -LiteralPath $PluginAssembly -Destination $outputDirectory -Force
Copy-Item -LiteralPath $sqliteAssembly -Destination $outputDirectory -Force
Copy-Item -LiteralPath $NativeSqliteAssembly -Destination (Join-Path $outputDirectory "e_sqlite3.dll") -Force
& $testExe
if ($LASTEXITCODE -ne 0) { throw "Department prompt tests failed." }
