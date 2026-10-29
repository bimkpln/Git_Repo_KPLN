$ErrorActionPreference = 'Stop'
$libraryProject = Join-Path $PSScriptRoot '../KPLN_Library_Schedules'
$libraryDll = Join-Path $libraryProject 'bin/Debug/2024/KPLN_Library_Schedules.dll'
$formsDll = Join-Path $libraryProject '../../KPLN_Library_Forms/KPLN_Library_Forms/bin/Debug/2024/KPLN_Library_Forms.dll'
$revitDll = 'C:/Program Files/Autodesk/Revit 2024/RevitAPI.dll'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
[void][Reflection.Assembly]::LoadFrom($revitDll)
[void][Reflection.Assembly]::LoadFrom($formsDll)
[void][Reflection.Assembly]::LoadFrom($libraryDll)
$refs = @($revitDll, $formsDll, $libraryDll, [Windows.Window].Assembly.Location,
    [Windows.Media.Visual].Assembly.Location, [Windows.DependencyObject].Assembly.Location,
    'System.Xml.Linq', 'System.Xml', 'System.Core', 'System.Xaml')
Add-Type -Path (Join-Path $PSScriptRoot 'Verify.cs') -ReferencedAssemblies $refs
$output = Join-Path $PSScriptRoot 'output'
[void](New-Item -ItemType Directory -Path $output -Force)
[ScheduleLibraryVerification]::Run($output)
