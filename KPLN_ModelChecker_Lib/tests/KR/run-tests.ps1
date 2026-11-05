param([string]$MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe")
$ErrorActionPreference = "Stop"
$projects = @("Ifc\DecisionTests", "Ifc\MarkSourceTests", "Ifc\CountingTypeTests", "Ifc\MaterialCodeTests", "Concrete\Tests")
$executables = @("Ifc\bin\DecisionTests.exe", "Ifc\bin\mark-source\MarkSourceTests.exe", "Ifc\bin\counting-type\CountingTypeTests.exe", "Ifc\bin\material-code\MaterialCodeTests.exe", "Concrete\bin\Tests.exe")
foreach ($project in $projects) {
    & $MSBuild (Join-Path $PSScriptRoot "$project.csproj") /t:Build /v:quiet /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $project" }
}
foreach ($executable in $executables) {
    & (Join-Path $PSScriptRoot $executable)
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $executable" }
}
