param(
    [Parameter(Mandatory=$true)][string]$Workspace,
    [string]$PythonHome = $env:MONOLITH_PYTHON_HOME,
    [Parameter(ValueFromRemainingArguments=$true)][string[]]$PythonArguments
)
$ErrorActionPreference = 'Stop'
$workspacePath = (Resolve-Path -LiteralPath $Workspace).ProviderPath
if ($workspacePath.StartsWith('\\')) { throw 'Use a local working copy, not the shared network folder.' }
if ([string]::IsNullOrWhiteSpace($PythonHome)) {
    $PythonHome = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python'
}
$pythonRoot = (Resolve-Path -LiteralPath $PythonHome).ProviderPath
if (-not (Test-Path -LiteralPath (Join-Path $pythonRoot 'python312.dll'))) {
    throw 'CPython 3.12 x64 is required. Discover the local runtime and pass -PythonHome; do not use another user profile.'
}
$runner = Join-Path $workspacePath 'iteration3\PythonRunner.exe'
if (-not (Test-Path -LiteralPath $runner)) { throw 'Restore the research code checkpoint first.' }
$previousPythonHome = $env:MONOLITH_PYTHON_HOME
$previousPythonPath = $env:PYTHONPATH
try {
    $env:MONOLITH_PYTHON_HOME = $pythonRoot
    $env:PYTHONPATH = (Join-Path $workspacePath '.geometry_deps')
    Push-Location -LiteralPath $workspacePath
    try { & $runner -B -X utf8 @PythonArguments; $resultCode=$LASTEXITCODE }
    finally { Pop-Location }
} finally {
    $env:MONOLITH_PYTHON_HOME = $previousPythonHome
    $env:PYTHONPATH = $previousPythonPath
}
exit $resultCode
