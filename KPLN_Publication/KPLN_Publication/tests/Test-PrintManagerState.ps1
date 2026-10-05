param([Parameter(Mandatory=$true)][string]$SourcePath)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath $SourcePath -Raw
$start = $source.IndexOf('        private static void ConfigureSingleSheet(')
$end = $source.IndexOf('        private static Action<PrintParameters> CaptureParameters(', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Print manager methods not found' }
$methods = $source.Substring($start, $end - $start)
$prefix = @'
using System;
using System.Collections.Generic;
public enum PrintRange { Current, Visible, Select }
public enum VirtualPrinterType { None, AdobePDF }
public class PrintManager
{
    public VirtualPrinterType IsVirtual;
    public PrintRange PrintRange;
    public bool SeedToFile, SeedCombined;
    public int ToFileWrites;
    public bool PrintToFile
    {
        get { return SeedToFile; }
        set
        {
            if (!value && IsVirtual != VirtualPrinterType.None)
                throw new InvalidOperationException("The PrintToFile property cannot be set to false while the printer is Virtual.");
            ToFileWrites++; SeedToFile = value;
        }
    }
    public int CombinedWrites;
    public bool CombinedFile
    {
        get { if (IsVirtual == VirtualPrinterType.None) throw new InvalidOperationException("Not virtual"); return SeedCombined; }
        set
        {
            if (IsVirtual == VirtualPrinterType.None) throw new InvalidOperationException("Not virtual");
            if (!value && PrintRange != PrintRange.Select) throw new InvalidOperationException("CombinedFile property cannot be set to false when the Print Range is Current/Visible!");
            CombinedWrites++; SeedCombined = value;
        }
    }
}
public static class ManagerTests
{
'@
$suffix = @'
    public static string Run()
    {
        int count = 0;
        foreach (VirtualPrinterType printer in new[] { VirtualPrinterType.None, VirtualPrinterType.AdobePDF })
        foreach (PrintRange range in new[] { PrintRange.Current, PrintRange.Visible, PrintRange.Select })
        foreach (bool combined in new[] { false, true })
        {
            var manager = new PrintManager { IsVirtual = printer, PrintRange = range, SeedCombined = combined };
            bool? original = manager.IsVirtual != VirtualPrinterType.None ? (bool?)manager.CombinedFile : null;
            ConfigureSingleSheet(manager);
            if (manager.PrintRange != PrintRange.Current || !manager.PrintToFile) throw new Exception("Wrong single-sheet configuration");
            if (printer != VirtualPrinterType.None && !manager.CombinedFile) throw new Exception("Illegal Current configuration");
            RestoreRangeAndCombined(manager, range, original);
            if (manager.PrintRange != range) throw new Exception("Wrong restored range");
            if (printer != VirtualPrinterType.None && manager.CombinedFile != combined) throw new Exception("Combined not restored");
            if (printer == VirtualPrinterType.None && manager.CombinedWrites != 0) throw new Exception("Physical printer touched");
            count++;
        }
        int restorationCases = 0;
        foreach (VirtualPrinterType printer in new[] { VirtualPrinterType.None, VirtualPrinterType.AdobePDF })
        foreach (bool original in new[] { false, true })
        foreach (bool current in new[] { false, true })
        {
            var manager = new PrintManager { IsVirtual = printer, SeedToFile = current };
            var result = new Dictionary<string, object>();
            RestorePrintToFile(manager, original, result);
            bool normalized = printer != VirtualPrinterType.None && !original && current;
            if (manager.PrintToFile != (normalized || original)) throw new Exception("Wrong restored PrintToFile");
            if (result.ContainsKey("restoration_warnings") != normalized) throw new Exception("Missing normalization warning");
            if (current == original && manager.ToFileWrites != 0) throw new Exception("Unnecessary setter call");
            restorationCases++;
        }
        var errors = new List<string>();
        bool continued = false;
        RestoreStep(() => { throw new InvalidOperationException("fixture failure"); }, "First", errors);
        RestoreStep(() => continued = true, "Second", errors);
        if (!continued || errors.Count != 1) throw new Exception("Recovery stopped or swallowed error");
        return "PASS: " + count + " printer/range/combined round trips, " + restorationCases + " PrintToFile restoration cases and independent recovery after failure.";
    }
}
'@
Add-Type -TypeDefinition ($prefix + $methods + $suffix)
[ManagerTests]::Run()
