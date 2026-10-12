param([string]$Compiler = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe')
$ErrorActionPreference = 'Stop'
$project = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$temp = Join-Path ([IO.Path]::GetTempPath()) ('KPLN-VentTests-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp) | Out-Null
try {
    $sources = @(Get-ChildItem -LiteralPath (Join-Path $project 'Forms\Entities\VentConfigurator') -Filter '*.cs' |
        Where-Object Name -ne 'OpenDocumentItem.cs' | ForEach-Object FullName)
    $sources += Join-Path $project 'Forms\Entities\OV_VentConfiguratorM.cs'
    foreach ($name in @('VentConfiguratorSettings', 'VentTypeName', 'ProjectPathMatcher')) {
        $sources += Join-Path $project "Services\VentConfigurator\$name.cs"
    }
    $sources += Join-Path $PSScriptRoot 'DomainTests.cs'
    $exe = Join-Path $temp 'DomainTests.exe'
    & $Compiler /nologo /target:exe /langversion:7.3 /nowarn:0649 /r:System.Core.dll "/out:$exe" $sources
    if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw 'Domain checks failed.' }
} finally {
    $resolved = [IO.Path]::GetFullPath($temp)
    $root = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (!$resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid cleanup directory.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
