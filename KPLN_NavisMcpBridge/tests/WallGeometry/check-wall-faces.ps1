param([Parameter(Mandatory=$true)][string]$AssemblyPath)
$ErrorActionPreference = 'Stop'
$navisDir = 'C:\Program Files\Autodesk\Navisworks Manage 2020'
foreach ($name in @('Autodesk.Navisworks.Interop.ComApi.dll','Autodesk.Navisworks.Api.dll','Autodesk.Navisworks.ComApi.dll','Autodesk.Navisworks.Clash.dll')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path $navisDir $name))
}
$assembly = [Reflection.Assembly]::LoadFrom($AssemblyPath)
$pointType = $assembly.GetType('KPLN_NavisMcpBridge.Point3Dto')
$diagnosticType = $assembly.GetType('KPLN_NavisMcpBridge.WallGeometryDiagnostics')
$service = $assembly.GetType('KPLN_NavisMcpBridge.ClashService')
$fit = $service.GetMethod('DescribeCompleteWallMesh',[Reflection.BindingFlags]'NonPublic,Static')
$local = $service.GetMethod('DescribeLocalWallFaces',[Reflection.BindingFlags]'NonPublic,Static')
$listType = [System.Collections.Generic.List``1].MakeGenericType($pointType)
function Point($x,$y,$z) {
    $p=[Activator]::CreateInstance($pointType); $p.X=$x; $p.Y=$y; $p.Z=$z; return $p
}
function Wall([double]$length=10, [double]$thickness=.25, [double]$height=9) {
    $a=$length/2; $b=$thickness/2
    $p=@((Point (-$a) (-$b) 0),(Point $a (-$b) 0),(Point $a $b 0),(Point (-$a) $b 0),
         (Point (-$a) (-$b) $height),(Point $a (-$b) $height),(Point $a $b $height),(Point (-$a) $b $height))
    $mesh=[Activator]::CreateInstance($listType)
    foreach ($face in @(@(0,2,1),@(0,3,2),@(4,5,6),@(4,6,7),@(0,1,5),@(0,5,4),
                       @(1,2,6),@(1,6,5),@(2,3,7),@(2,7,6),@(3,0,4),@(3,4,7))) {
        foreach($i in $face) { $mesh.Add($p[$i]) }
    }
    return ,$mesh
}
function Invoke-WallMeasurement($mesh,$contact) {
    $d=[Activator]::CreateInstance($diagnosticType)
    $analysis=$fit.Invoke($null,@($mesh,$d))
    return $local.Invoke($null,@($analysis,$contact))
}
function Check($condition,$message) { if(-not $condition) { throw $message } }
$mesh=Wall
$r=Invoke-WallMeasurement $mesh (Point 0 .125 4)
Check ($r.Usable -and [Math]::Abs($r.HalfThickness-.125) -lt 1e-8 -and [Math]::Abs($r.NearestEndFace-5) -lt 1e-8) 'full wall broad/end faces'
$r=Invoke-WallMeasurement $mesh (Point 5 0 4)
Check ($r.Usable -and $r.NearestEndFace -lt 1e-8) 'end contact'
$angle=-.8; $c=[Math]::Cos($angle); $s=[Math]::Sin($angle)
$rotated=[Activator]::CreateInstance($listType)
foreach($p in $mesh) { $rotated.Add((Point ($p.X*$c-$p.Y*$s) ($p.X*$s+$p.Y*$c) $p.Z)) }
$r=Invoke-WallMeasurement $rotated (Point (-.125*$s) (.125*$c) 4)
Check ($r.Usable -and [Math]::Abs($r.AxisX*$c+$r.AxisY*$s) -gt .999999 -and $r.NearestBroadFace -lt 1e-7) 'negative diagonal wall'
$shifted=[Activator]::CreateInstance($listType)
foreach($p in $rotated) { $shifted.Add((Point ($p.X+7212103) ($p.Y+1041196) ($p.Z+48))) }
$r=Invoke-WallMeasurement $shifted (Point (7212103-.125*$s) (1041196+.125*$c) 52)
Check ($r.Usable -and [Math]::Abs($r.HalfThickness-.125) -lt 1e-7) 'large world coordinates'
$sheet=[Activator]::CreateInstance($listType)
foreach($i in 12..17) { $sheet.Add($mesh[$i]) }
$r=Invoke-WallMeasurement $sheet (Point 0 (-.125) 4)
Check (-not $r.Usable -and $r.Reason -eq 'wall-zero-thickness-surface') 'material surface cannot impersonate whole wall'
$open=[Activator]::CreateInstance($listType)
foreach($i in 0..35) { if($i -lt 24 -or $i -gt 29) { $open.Add($mesh[$i]) } }
$r=Invoke-WallMeasurement $open (Point 0 (-.125) 4)
Check (-not $r.Usable -and $r.Reason -eq 'wall-opposite-broad-faces-missing') 'missing opposing broad surface'
$square=Wall 1 1 9
$r=Invoke-WallMeasurement $square (Point 0 .5 4)
Check (-not $r.Usable -and $r.Reason -eq 'wall-broad-direction-ambiguous') 'ambiguous broad direction'
$mesh.Add((Point 1 (-.125) 0)); $mesh.Add((Point 1 .125 0)); $mesh.Add((Point 1 .125 7))
$mesh.Add((Point 1 (-.125) 0)); $mesh.Add((Point 1 .125 7)); $mesh.Add((Point 1 (-.125) 7))
$r=Invoke-WallMeasurement $mesh (Point 1 .125 4)
Check ($r.Usable -and $r.NearestEndFace -lt 1e-8) 'internal reveal contact'
$layers=Wall 10 .25 9
$inner=Wall 10 .15 9
foreach($p in $inner) { $layers.Add($p) }
$r=Invoke-WallMeasurement $layers (Point 0 .125 4)
Check ($r.Usable -and [Math]::Abs($r.HalfThickness-.125) -lt 1e-8) 'multiple layers use complete outer thickness'
Write-Output '9 wall geometry checks passed against built add-in.'
