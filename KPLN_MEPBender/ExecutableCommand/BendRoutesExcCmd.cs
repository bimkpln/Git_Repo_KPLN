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
                    _entity.AngleDegrees,
                    _entity.GetSelectedDirections(),
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
                    result.Message);

                if (_entity.AutoClearObstaclesAfterRun)
                    _entity.ClearObstacles();

                if (_entity.AutoClearRoutesAfterRun)
                    _entity.ClearRoutes();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("MEP Bender", ex.ToString());
                _entity.SetStatus("Ошибка при выполнении MEP Bender.");
                return Result.Failed;
            }
        }
    }
}
