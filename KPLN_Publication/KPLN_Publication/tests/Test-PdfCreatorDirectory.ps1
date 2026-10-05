param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$PdfCreatorDirectory = 'C:\Program Files\PDFCreator'
)
$ErrorActionPreference = 'Stop'
# Загружаем реальные библиотеки установленного PDFCreator. Реестр и очередь не меняем.
$null = [Reflection.Assembly]::LoadFrom((Join-Path $PdfCreatorDirectory 'DataStorage.dll'))
$null = [Reflection.Assembly]::LoadFrom((Join-Path $PdfCreatorDirectory 'PDFCreator.Settings.dll'))
Add-Type -TypeDefinition (Get-Content -LiteralPath $SourcePath -Raw)
$scopeType = [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetType('KPLN_Publication.PdfCreatorSettingsScope') } | Where-Object { $null -ne $_ } | Select-Object -First 1
$encoder = $scopeType.GetMethod('EncodeDirectory', [Reflection.BindingFlags]'Static,NonPublic')
if (!$encoder) { throw 'Production encoder not found' }
$paths = @(
    'C:\PDF_Print\20261002T090650Z-3d9aa00f0fd1423d9851cf90b3c6bb21',
    'C:\PDF_Print\АР эталон',
    'C:\PDF_Print\new\test',
    'D:\Проверка (АР)\Листы 1-4'
)
foreach ($path in $paths) {
    $encoded = $encoder.Invoke($null, @($path))
    if ($encoded -cne [pdfforge.DataStorage.Data]::EscapeString($path)) { throw "Serialization differs: $path" }
    $data = [pdfforge.DataStorage.Data]::CreateDataStorage()
    $data.SetValue('TargetDirectory', $encoded)
    $profile = [pdfforge.PDFCreator.Conversion.Settings.ConversionProfile]::new()
    $profile.ReadValues($data, '')
    if ($profile.TargetDirectory -cne $path) { throw "Directory changed after real profile load: $path" }
    if (![IO.Path]::IsPathRooted($profile.TargetDirectory)) { throw 'Not rooted' }
}
# Негативный контроль: старое значение действительно теряет каталог при загрузке профиля.
$data = [pdfforge.DataStorage.Data]::CreateDataStorage()
$data.SetValue('TargetDirectory', $paths[0])
$profile = [pdfforge.PDFCreator.Conversion.Settings.ConversionProfile]::new()
$profile.ReadValues($data, '')
if (![string]::IsNullOrEmpty($profile.TargetDirectory)) { throw 'Old bug was not reproduced' }
"PASS: $($paths.Count) production-encoder round trips through installed PDFCreator; old raw-path bug reproduced. No registry writes or print jobs."
