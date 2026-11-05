using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Core;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckIfcRebarMarks : AbstrKRCheck
    {
        public const string CheckId = Common.KR.Ifc.Checker.CheckId;
        public CheckIfcRebarMarks() { PluginName = "КР: Проверка марок IFC"; }
        protected override IDictionary<string, object> Evaluate(Document doc, Element[] candidates)
            => Common.KR.Ifc.Checker.Run(doc, candidates);
        public static string RunJson(UIApplication app) => RunJson<CheckIfcRebarMarks>(app);
    }
}
