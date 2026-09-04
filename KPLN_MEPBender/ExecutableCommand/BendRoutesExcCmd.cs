using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_MEPBender.Forms.Entities;
using KPLN_MEPBender.Services.Clashes;
using KPLN_MEPBender.Services.Routing;
using System;

namespace KPLN_MEPBender.ExecutableCommand
{
    public sealed class BendRoutesExcCmd : IExecutableCommand
    {
        private readonly MepBenderM _entity;

        public BendRoutesExcCmd(MepBenderM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            Document doc = uiDoc.Document;
            try
            {
                MepBendRequest request = new MepBendRequest(
                    doc,
                    doc.ActiveView,
                    _entity.RouteElementIds,
                    _entity.ObstacleReferences,
                    _entity.OffsetMm,
                    _entity.OffsetIterationStepMm,
                    _entity.AngleDegrees,
                    _entity.GetSelectedDirections(),
                    _entity.AlignVerticalBendByLowest,
                    _entity.AnalyzeCollisions);

                MepBendResult result = new MepRouteBender().Execute(request);

                if (_entity.AnalyzeCollisions)
                {
                    new IosClasherService().Analyze(new ClashAnalyzeRequest(
                        doc,
                        result.CreatedElementIds,
                        _entity.AnalyzeCollisions));
                }

                _entity.SetStatus(
                    result.GeometryWasChanged ? null : "Геометрия трасс пока не изменялась.",
                    result.Message,
                    GetResultHelpForeground(result));

                if (result.HasInvalidFittingFamilyFailure)
                    ShowInvalidFittingFamilyDialog();

                if (_entity.AutoClearObstaclesAfterRun)
                    _entity.ClearObstacles();

                if (_entity.AutoClearRoutesAfterRun)
                    _entity.ClearRoutes();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("MEP Bender", ex.ToString());
                _entity.SetStatus("Ошибка при выполнении MEP Bender.", null, "#FF5C5C");
                return Result.Failed;
            }
        }

        private string GetResultHelpForeground(MepBendResult result)
        {
            if (result == null || !result.GeometryWasChanged || result.ProcessedRouteIds.Count == 0)
                return "#FF5C5C";

            bool hasErrors = result.SkippedRouteCount > 0
                             || result.FailedRouteCount > 0
                             || result.FittingFailureCount > 0
                             || result.ReconnectFailureCount > 0
                             || result.HasInvalidFittingFamilyFailure
                             || result.HasInsufficientSpaceFailure;

            return hasErrors ? "#FFD166" : "#6FD37A";
        }
        private void ShowInvalidFittingFamilyDialog()
        {
            TaskDialog dialog = new TaskDialog("MEP Bender")
            {
                MainInstruction = "Проверь семейства фасонных элементов",
                MainContent = "Revit не нашёл подходящий отвод/соединительную деталь для выбранного угла. Проверь семейства, таблицы углов и тип трассы.",
                CommonButtons = TaskDialogCommonButtons.Ok
            };
            dialog.Show();
        }
    }
}
