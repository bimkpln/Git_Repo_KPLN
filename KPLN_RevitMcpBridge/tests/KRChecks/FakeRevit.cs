using System;
using System.Linq;
using KPLN_RevitMcpBridge.Server;
namespace Autodesk.Revit.DB
{
    public class Document { public bool IsFamilyDocument; public bool IsLinked; }
}
namespace Autodesk.Revit.UI
{
    public class UIDocument { public Autodesk.Revit.DB.Document Document = new Autodesk.Revit.DB.Document(); }
    public class UIApplication { public UIDocument ActiveUIDocument = new UIDocument(); }
}
namespace KPLN_ModelChecker_User.Checks.KR
{
    internal static class KRCheckAccess
    {
        internal static bool Allowed = true;
        internal static bool Throws;
        internal static bool IsAllowed() {
            if (Throws) throw new Exception("DB unavailable");
            return Allowed;
        }
    }
    internal static class FakeReport
    {
        internal static int Calls;
        internal static string Override;
        internal static bool Throws;
        internal static string Run(string id) {
            Calls++;
            if (Throws) throw new InvalidOperationException("algorithm failed");
            return Override ?? Json.Serialize(new {
                check_id = id, status = "partial", error_count = 201, review_count = 1,
                notes = new[] { "closed worksets" },
                issues = Enumerable.Range(0, 201).Select(n => new { element_id = n.ToString(), status = "mismatch" }).ToArray(),
                skipped_elements = new[] { new { element_id = "skip", message = "parameter" } },
                run_warnings = new[] { new { message = "materials" } }
            });
        }
    }
}
namespace KPLN_ModelChecker_Lib.Commands
{
    public static class CheckIfcRebarMarks
    {
        public const string CheckId = "CheckIfcRebarMarks";
        public static string RunJson(Autodesk.Revit.UI.UIApplication app) { return KPLN_ModelChecker_User.Checks.KR.FakeReport.Run(CheckId); }
    }
}
namespace KPLN_ModelChecker_Lib.Commands
{
    public static class CheckRebarConcreteClass
    {
        public const string CheckId = "CheckRebarConcreteClass";
        public static string RunJson(Autodesk.Revit.UI.UIApplication app) { return KPLN_ModelChecker_User.Checks.KR.FakeReport.Run(CheckId); }
    }
}
