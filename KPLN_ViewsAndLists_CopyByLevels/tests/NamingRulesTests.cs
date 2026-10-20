using KPLN_ViewsAndLists_CopyByLevels.Common;
using System;
using System.Linq;

internal static class NamingRulesTests
{
    private static int _count;

    private static void Equal<T>(T actual, T expected, string scenario)
    {
        if (!Equals(actual, expected)) throw new Exception(scenario + ": expected " + expected + ", actual " + actual);
        _count++;
    }

    private static void Reject(Action action, string scenario)
    {
        try { action(); }
        catch (ArgumentException) { _count++; return; }
        throw new Exception(scenario + ": expected rejection");
    }

    public static int Main()
    {
        Equal(string.Join(",", NamingRules.SelectLevels(new[] { 0.6, 1.5, 1.501, 3.0, 4.49, 4.5 }, 0, 1.5).Select(b => b ? "1" : "0")),
            "0,1,0,1,0,1", "Technical levels, equal elevations and inclusive threshold");
        Equal(string.Join(",", NamingRules.SelectLevels(new[] { 6.0, 3.0, -1.0, 0.0 }, 0, 1.5).Select(b => b ? "1" : "0")),
            "1,1,0,0", "Input order does not determine selection");
        Reject(() => NamingRules.SelectLevels(new[] { 1.0 }, 0, 0), "Zero gap");
        Reject(() => NamingRules.SelectLevels(new[] { 1.0 }, 0, double.NaN), "NaN gap");
        Equal(NamingRules.Elevation(4.95), "+4.950", "Positive elevation");
        Equal(NamingRules.Elevation(-2.535), "-2.535", "Basement elevation");
        Equal(NamingRules.Expand("1.1.{Этаж}.", "6", "Этаж 6", "+18.150", "", "1.1.1.", 2), "1.1.6.", "Room prefix mask");
        Equal(NamingRules.Expand("{Исходное} — {Уровень}", "2", "Этаж 2", "+4.950", "", "Текст {не маска}", 1),
            "Текст {не маска} — Этаж 2", "Source text is not reinterpreted as a mask");
        Reject(() => NamingRules.Expand("{Etazh}", "2", "", "", "", "", 1), "Unknown token");
        Equal(NamingRules.GuessFloor("С1.1_Этаж_02_+4.950", 99), "2", "MTRS floor name");
        Equal(NamingRules.GuessFloor("Подвал", -1), "-1", "Unknown floor keeps explicit fallback");
        Console.WriteLine("PASS: " + _count + " meaningful naming/selection cases");
        return 0;
    }
}
