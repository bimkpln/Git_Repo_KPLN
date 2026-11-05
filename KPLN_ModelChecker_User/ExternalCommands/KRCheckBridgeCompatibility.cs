using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;

// Compatibility entry points used by the existing Codex bridge via reflection.
// Algorithms live exclusively in KPLN_ModelChecker_Lib.
namespace KPLN_ModelChecker_User.Checks.KR.Ifc
{
    public static class Checker
    {
        public const string CheckId = CheckIfcRebarMarks.CheckId;
        public static string RunJson(UIApplication app) => CheckIfcRebarMarks.RunJson(app);
    }
}

namespace KPLN_ModelChecker_User.Checks.KR.Concrete
{
    public static class Entry
    {
        public const string CheckId = CheckRebarConcreteClass.CheckId;
        public static string RunJson(UIApplication app) => CheckRebarConcreteClass.RunJson(app);
    }
}
