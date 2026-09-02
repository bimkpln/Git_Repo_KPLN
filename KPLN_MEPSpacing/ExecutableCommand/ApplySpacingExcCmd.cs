using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_MEPSpacing.Common;
using KPLN_MEPSpacing.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;

namespace KPLN_MEPSpacing.ExecutableCommand
{
    public sealed class ApplySpacingExcCmd : IExecutableCommand
    {
        private readonly MepSpacingM _entity;

        public ApplySpacingExcCmd(MepSpacingM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            if (!_entity.TryGetDistanceMm(out double distanceMm) || distanceMm <= 0)
            {
                _entity.SetErrorStatus("Укажи положительное расстояние в мм.");
                return Result.Cancelled;
            }

            try
            {
                SpacingApplyResult result;
                using (Transaction transaction = new Transaction(uiDoc.Document, "KPLN: Ровный шаг MEP"))
                {
                    transaction.Start();
                    result = MepSpacingService.ApplySpacing(uiDoc.Document, _entity.GetAllElementIds(), _entity.BaseElementIds, distanceMm, _entity.CalculationMode);
                    transaction.Commit();
                }

                _entity.SetResultStatus(result);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                _entity.SetErrorStatus(ex.Message);
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }
    }
}
