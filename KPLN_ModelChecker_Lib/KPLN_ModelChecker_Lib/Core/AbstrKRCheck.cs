using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands.Common.KR;

namespace KPLN_ModelChecker_Lib.Core
{
    /// <summary>Read-only KR checks using the standard AbstrCheck execution stages.</summary>
    public abstract class AbstrKRCheck : AbstrCheck
    {
        public KRCheckResult Report { get; private set; }
        protected override bool ExecuteOnEmptyCollection => true;
        protected override bool UseDefaultErrorReporting => false;

        public override Element[] GetElemsToCheck()
        {
            RequireDocument();
            using (var collector = new FilteredElementCollector(CheckDocument))
                return collector.OfCategory(BuiltInCategory.OST_Rebar)
                    .WhereElementIsNotElementType().ToArray();
        }

        protected override void ResetExecutionState()
        {
            Report = null;
            _checkerEntitiesCollHeap.Clear();
            CheckerEntitiesColl = new CheckerEntity[0];
            PrepareElemsErrorColl.Clear();
            RunElemsWarningColl.Clear();
            PreparedElemColl = null;
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            RequireDocument();
            // Filtering, uncertainty handling and geometry rules stay in the evaluator.
            // Do not skip an empty collection: material/workset limitations still matter.
            Report = new KRCheckResult(Evaluate(CheckDocument, elemColl));
            _checkerEntitiesCollHeap.AddRange(Report.ToCheckerEntities(CheckDocument));
            return Report.Status == "empty" ? CheckResultStatus.NoItemsToCheck : CheckResultStatus.Succeeded;
        }

        private void RequireDocument()
        {
            if (CheckDocument == null || CheckDocument.IsFamilyDocument)
                throw new InvalidOperationException("Откройте проект RVT.");
        }

        protected abstract IDictionary<string, object> Evaluate(Document doc, Element[] candidates);

        // Stable serialization boundary for the existing MCP bridge. No windows or model writes.
        protected static string RunJson<T>(UIApplication app) where T : AbstrKRCheck, new()
        {
            var check = new T();
            check.Set_UIAppData(app, app.ActiveUIDocument?.Document);
            check.ExecuteCheck(check.GetElemsToCheck(), false);
            return check.Report.ToJson();
        }
    }
}
