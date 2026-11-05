using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Core;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckRebarConcreteClass : AbstrKRCheck
    {
        public const string CheckId = Common.KR.Concrete.Entry.CheckId;
        public CheckRebarConcreteClass() { PluginName = "КР: Проверка класса бетона арматуры"; }
        protected override IDictionary<string, object> Evaluate(Document doc, Element[] candidates)
            => Common.KR.Concrete.Entry.Run(doc, candidates);
        public static string RunJson(UIApplication app) => RunJson<CheckRebarConcreteClass>(app);
    }
}
