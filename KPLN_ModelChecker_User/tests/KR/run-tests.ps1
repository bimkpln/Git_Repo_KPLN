param([string]$MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe", [string]$PreviewPath = "")
$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "..\..\..\KPLN_ModelChecker_Lib\tests\KR\run-tests.ps1") -MSBuild $MSBuild
$projects = @("Report\ReportTests", "Lifecycle\Lifecycle")
$executables = @("Report\bin\ReportTests.exe", "Lifecycle\bin\Lifecycle.exe")
foreach ($project in $projects) {
    & $MSBuild (Join-Path $PSScriptRoot "$project.csproj") /t:Build /v:quiet /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $project" }
}
foreach ($executable in $executables) {
    $program = Join-Path $PSScriptRoot $executable
    if ($executable -eq "Report\bin\ReportTests.exe" -and $PreviewPath) { & $program $PreviewPath } else { & $program }
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $executable" }
}
