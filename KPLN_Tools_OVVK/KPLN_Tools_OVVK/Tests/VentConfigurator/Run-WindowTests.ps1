param(
    [string]$Configuration = 'Debug2023',
    [string]$ProjectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')),
    [string]$PreviewPath
)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with Windows PowerShell -STA.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
$year = $Configuration -replace '\D', ''
$repo = [IO.Path]::GetFullPath((Join-Path $ProjectRoot '..\..'))
foreach ($dll in @('RevitAPI', 'RevitAPIUI')) {
    [Reflection.Assembly]::LoadFrom("C:\Program Files\Autodesk\Revit $year\$dll.dll") | Out-Null
}
[Reflection.Assembly]::LoadFrom("$repo\01_KPLN_Library\KPLN_Library_Forms\KPLN_Library_Forms\bin\Debug\$year\KPLN_Library_Forms.dll") | Out-Null
$assembly = [Reflection.Assembly]::LoadFrom("$ProjectRoot\bin\Debug\$year\KPLN_Tools_OVVK.dll")
$flags = [Reflection.BindingFlags]'Instance, Static, Public, NonPublic'
$prefix = 'KPLN_Tools_OVVK.Forms.Entities.VentConfigurator.'
$script:checks = 0
function Check($condition, $name) {
    if (!$condition) { throw $name }
    $script:checks++
}
function Get-Value($obj, $name) {
    $prop = $obj.GetType().GetProperty($name, $flags)
    if ($prop) { return ,$prop.GetValue($obj, $null) }
    return ,$obj.GetType().GetField($name, $flags).GetValue($obj)
}
function Set-Value($obj, $name, $value) {
    $prop = $obj.GetType().GetProperty($name, $flags)
    if ($prop) { $prop.SetValue($obj, $value, $null); return }
    $obj.GetType().GetField($name, $flags).SetValue($obj, $value)
}
function Invoke-Method($obj, $name, [object[]]$arguments = @()) {
    return ,$obj.GetType().GetMethod($name, $flags).Invoke($obj, $arguments)
}
function New-Entity($name, [object[]]$arguments = @()) {
    return ,[Activator]::CreateInstance($assembly.GetType($prefix + $name), $flags, $null, $arguments, $null)
}
function Pump {
    [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle)
}
function Command($vm, $name) { return Get-Value $vm $name }

$application = New-Object Windows.Application
$application.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
$bindingLog = New-Object IO.StringWriter
$listener = New-Object Diagnostics.TextWriterTraceListener($bindingLog)
[Diagnostics.PresentationTraceSources]::DataBindingSource.Listeners.Add($listener) | Out-Null
[Diagnostics.PresentationTraceSources]::DataBindingSource.Switch.Level = [Diagnostics.SourceLevels]::Warning
$window = $assembly.CreateInstance('KPLN_Tools_OVVK.Forms.OV_VentConfiguratorForm')
try {
    $vm = Get-Value $window 'CurrentOV_VentConfiguratorVM'
    $model = Get-Value $vm 'CurrentOV_VentConfiguratorM'
    $window.WindowStartupLocation = [Windows.WindowStartupLocation]::Manual
    $window.ShowInTaskbar = $false
    $window.Left = -20000
    $window.Top = -20000
    $window.Show()
    Pump
    Check ($window.FindResource('KPLN_Button') -ne $null) 'Shared style loaded'
    Check (!(Command $vm 'SaveCmd').CanExecute($null)) 'Save disabled without a type'
    Check (!(Get-Value $window 'CopyTypeButton').IsEnabled) 'Copy disabled without a type'

    $catalog = New-Entity 'SectionCatalog'
    Set-Value $catalog 'IntermediateCount' 2
    Set-Value $catalog 'InstallationWidthMm' ([double]1000)
    Set-Value $catalog 'InstallationHeightMm' ([double]1200)
    Set-Value $catalog 'FrameHeightMm' ([double]100)
    $choice = New-Entity 'SectionTypeChoice' @('Вентилятор', 'Вентилятор')
    for ($slot = 1; $slot -le 12; $slot++) {
        $definition = New-Entity 'SectionDefinition'
        $names = $assembly.GetType($prefix + 'InstallationConfiguration')
        $parameter = $names.GetMethod('ParameterName', $flags).Invoke($null, @($slot))
        Set-Value $definition 'ParameterName' $parameter
        Set-Value $definition 'DisplayName' ("Секция $slot")
        $listType = $definition.GetType().GetField('Choices', $flags).FieldType
        $choices = [Activator]::CreateInstance($listType)
        $choices.Add($choice)
        Set-Value $definition 'Choices' $choices
        Set-Value $definition 'SourceType' $choice
        foreach ($dimension in @('DefaultLengthMm', 'DefaultWidthMm', 'DefaultHeightMm')) {
            Set-Value $definition $dimension ([double]200)
        }
        $catalog.Add($slot, $definition)
    }
    $package = New-Entity 'FamilyPackage'
    Set-Value $package 'Catalog' $catalog
    Invoke-Method $vm 'SetFamily' @($package) | Out-Null
    (Command $vm 'CreateTypeCmd').Execute($null)
    Pump
    Check ((Get-Value $model 'Types').Count -eq 1) 'Create command adds type'
    Check ((Get-Value $window 'CopyTypeButton').IsEnabled) 'Copy enabled after create'
    Check ((Get-Value $window 'SaveButton').IsEnabled) 'Save enabled after create'
    Set-Value $model 'WorkingFamilyPath' 'C:\test.rfa'
    Set-Value $model 'CanUpdate' $true
    Invoke-Method $model 'RefreshState' | Out-Null
    [Windows.Input.CommandManager]::InvalidateRequerySuggested()
    Pump
    Check ((Get-Value $window 'UpdateFamilyButton').IsEnabled) 'Update enabled for authorized user with family'
    (Command $vm 'CopyTypeCmd').Execute($null)
    Pump
    Check ((Get-Value $model 'Types').Count -eq 2) 'Copy command adds independent type'
    $current = Get-Value $model 'CurrentType'
    (Get-Value $window 'TypeNameTextBox').Text = 'П1_Проверка'
    Check ((Get-Value $current 'Name') -eq 'П1_Проверка') 'Type rename reaches model'
    Check (!(Get-Value $current 'HasAutomaticName')) 'Manual name stops automatic naming'

    Invoke-Method $vm 'SetBusy' @($true) | Out-Null
    Pump
    Check (!(Get-Value $window 'SaveButton').IsEnabled) 'Save disabled during request'
    Check (!(Get-Value $window 'EditorPanel').IsEnabled) 'Editor disabled during request'
    $window.Close()
    Check $window.IsVisible 'Closing blocked during request'
    Invoke-Method $vm 'SetBusy' @($false) | Out-Null
    Invoke-Method $vm 'SetStatus' @('Проверка сообщения', $true) | Out-Null
    Pump
    Check ((Get-Value $window 'ErrorDetailsButton').Visibility -eq 'Visible') 'Error details visible'
    Invoke-Method $vm 'SetStatus' @('Готово', $false) | Out-Null

    Invoke-Method $vm 'SwitchFamilyContext' @('one') | Out-Null
    Invoke-Method $vm 'SetFamily' @($package) | Out-Null
    (Command $vm 'CreateTypeCmd').Execute($null)
    $savedSelection = Get-Value $model 'CurrentType'
    Invoke-Method $vm 'SwitchFamilyContext' @('two') | Out-Null
    Check ((Get-Value $model 'Types').Count -eq 0) 'New context starts empty'
    Invoke-Method $vm 'SwitchFamilyContext' @('one') | Out-Null
    Check ([object]::ReferenceEquals((Get-Value $model 'CurrentType'), $savedSelection)) 'Context restores selection and edits'

    Set-Value $window 'SelectedTab' 1
    Pump
    $window.UpdateLayout()
    Check ((Get-Value $window 'SectionBlocksPanel').Children.Count -gt 0) 'Composition renders'
    $icons = Get-Value $window '_sectionIcons'
    Check ($icons.Count -gt 0) 'Section icons loaded'
    foreach ($icon in $icons.Values) { Check ($null -ne $icon) 'Icon resource resolves' }
    if ($PreviewPath) {
        $content = $window.Content
        $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap([int]$content.ActualWidth, [int]$content.ActualHeight, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
        $visual = New-Object Windows.Media.DrawingVisual
        $drawing = $visual.RenderOpen()
        $rect = New-Object Windows.Rect(0, 0, $content.ActualWidth, $content.ActualHeight)
        $drawing.DrawRectangle($window.Background, $null, $rect)
        $drawing.DrawRectangle((New-Object Windows.Media.VisualBrush($content)), $null, $rect)
        $drawing.Close()
        $bitmap.Render($visual)
        $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream = [IO.File]::Create($PreviewPath)
        try { $encoder.Save($stream) } finally { $stream.Dispose() }
    }

    $active = Get-Value $model 'Configuration'
    $section = (Get-Value $active 'Blocks')[0]
    $dimension = Get-Value $section 'Length'
    $info = Get-Value $active 'Info'
    Check ([string]::IsNullOrWhiteSpace($bindingLog.ToString())) ("WPF binding errors: " + $bindingLog.ToString())
    $window.Close()
    Check ((Get-Value $vm 'FamilyService').GetType().GetProperty('IsClosed', $flags).GetValue((Get-Value $vm 'FamilyService'), $null)) 'Service closed'
    foreach ($entity in @($section, $dimension, $info)) {
        $base = $entity.GetType().BaseType
        $subscribers = $base.GetField('PropertyChanged', $flags).GetValue($entity)
        $vmSubscriptions = @()
        if ($subscribers) {
            $vmSubscriptions = @($subscribers.GetInvocationList() | Where-Object { [object]::ReferenceEquals($_.Target, $vm) })
        }
        Check ($vmSubscriptions.Count -eq 0) 'VM subscriptions removed'
    }
    Write-Output "PASS: $script:checks WPF/VM checks."
} finally {
    Invoke-Method $window 'ReleaseResources' | Out-Null
    $application.Shutdown()
}
