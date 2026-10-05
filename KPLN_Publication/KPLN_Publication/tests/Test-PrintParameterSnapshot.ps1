param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [ValidateSet('2023','2024')][string]$RevitVersion = '2023'
)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath $SourcePath -Raw
$start = $source.IndexOf('        private static Action<PrintParameters> CaptureParameters(')
$end = $source.IndexOf('        private static void WaitForPdf(', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Snapshot method not found' }
$method = $source.Substring($start, $end - $start)
# Compile the actual production method against a strict API double, not a copied algorithm.
$prefix = @'
using System;
public enum ZoomType { Zoom, FitToPage }
public enum PaperPlacementType { Center, Margins }
public enum MarginType { PrinterLimit, UserDefined }
public class PrintParameters
{
    public bool HideCropBoundaries, HideReforWorkPlanes, HideScopeBoxes, HideUnreferencedViewTags;
    public ZoomType ZoomType;
    public PaperPlacementType PaperPlacement;
    public MarginType SeedMargin;
    public int MarginReads, MarginWrites;
    public MarginType MarginType
    {
        get { if (PaperPlacement != PaperPlacementType.Margins) throw new InvalidOperationException("The PaperPlacement is NOT Margins."); MarginReads++; return SeedMargin; }
        set { if (PaperPlacement != PaperPlacementType.Margins) throw new InvalidOperationException("The PaperPlacement is NOT Margins."); MarginWrites++; SeedMargin = value; }
    }
    public int RasterQuality, HiddenLineViews, ColorDepth, PaperSize, PageOrientation;
    public int OffsetReads, OffsetWrites;
    private double _x = -0.25, _y = 0.75;
    private void CheckOffsets()
    {
        if (PaperPlacement != PaperPlacementType.Margins || MarginType != MarginType.UserDefined)
            throw new InvalidOperationException("Current PaperPlacement is NOT Margins and Current MarginType is NOT User defined.");
    }
    public double UserDefinedMarginX { get { CheckOffsets(); OffsetReads++; return _x; } set { CheckOffsets(); OffsetWrites++; _x = value; } }
    public double UserDefinedMarginY { get { CheckOffsets(); OffsetReads++; return _y; } set { CheckOffsets(); OffsetWrites++; _y = value; } }
    public double OriginOffsetX { get { return UserDefinedMarginX; } set { UserDefinedMarginX = value; } }
    public double OriginOffsetY { get { return UserDefinedMarginY; } set { UserDefinedMarginY = value; } }
    public int ZoomReads, ZoomWrites;
    private int _zoom = 100;
    public int Zoom
    {
        get { ZoomReads++; if (ZoomType != ZoomType.Zoom) throw new InvalidOperationException("Current Zoom Type is NOT Zoom."); return _zoom; }
        set { ZoomWrites++; if (ZoomType != ZoomType.Zoom) throw new InvalidOperationException("Current Zoom Type is NOT Zoom."); _zoom = value; }
    }
}
public static class SnapshotTests
{
'@
$suffix = @'
    public static string Run()
    {
        int checks = 0;
        foreach (ZoomType original in new[] { ZoomType.FitToPage, ZoomType.Zoom })
        foreach (ZoomType targetMode in new[] { ZoomType.FitToPage, ZoomType.Zoom })
        foreach (int scale in new[] { 25, 100, 175 })
        foreach (PaperPlacementType placement in new[] { PaperPlacementType.Center, PaperPlacementType.Margins })
        foreach (MarginType margin in new[] { MarginType.PrinterLimit, MarginType.UserDefined })
        foreach (PaperPlacementType targetPlacement in new[] { PaperPlacementType.Center, PaperPlacementType.Margins })
        foreach (MarginType targetMargin in new[] { MarginType.PrinterLimit, MarginType.UserDefined })
        {
            var source = new PrintParameters { ZoomType = original, PaperPlacement = placement, SeedMargin = margin, HideCropBoundaries = true, PaperSize = 17 };
            if (original == ZoomType.Zoom) source.Zoom = scale;
            var restore = CaptureParameters(source);
            if (source.ZoomReads != (original == ZoomType.Zoom ? 1 : 0)) throw new Exception("Unexpected Zoom getter");
            var target = new PrintParameters { ZoomType = targetMode, PaperPlacement = targetPlacement, SeedMargin = targetMargin };
            restore(target);
            if (target.ZoomType != original) throw new Exception("Mode not restored");
            if (target.ZoomWrites != (original == ZoomType.Zoom ? 1 : 0)) throw new Exception("Unexpected Zoom setter");
            if (original == ZoomType.Zoom && target.Zoom != scale) throw new Exception("Scale not restored");
            if (!target.HideCropBoundaries || target.PaperSize != 17) throw new Exception("Other settings not restored");
            bool hasOffsets = placement == PaperPlacementType.Margins && margin == MarginType.UserDefined;
            if (source.OffsetReads != (hasOffsets ? 2 : 0) || target.OffsetWrites != (hasOffsets ? 2 : 0)) throw new Exception("Invalid offsets access");
            if (target.PaperPlacement != placement || (placement == PaperPlacementType.Margins && target.MarginType != margin)) throw new Exception("Placement not restored");
            if (placement == PaperPlacementType.Center && (source.MarginReads != 0 || target.MarginWrites != 0)) throw new Exception("Inactive margin accessed");
            if (hasOffsets && (target.UserDefinedMarginX != -0.25 || target.UserDefinedMarginY != 0.75)) throw new Exception("Offsets not restored");
            checks++;
        }
        return "PASS: " + checks + " snapshot cases (zoom, placement, margin, both target modes and offsets).";
    }
}
'@
$define = if ($RevitVersion -eq '2023') { "#define Revit2023`n" } else { '' }
Add-Type -TypeDefinition ($define + $prefix + $method + $suffix)
[SnapshotTests]::Run()
