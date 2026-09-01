using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_TrailingMEP.Common;
using KPLN_TrailingMEP.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_TrailingMEP.ExecutableCommand
{
    internal sealed class PickBundleExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;
        private readonly Window _window;

        public PickBundleExcCmd(RouteTraceM entity, Window window)
        {
            _entity = entity;
            _window = window;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            try
            {
                _window?.Hide();
                IList<Reference> references = uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    new MepCurveSelectionFilter(),
                    "Выбери трубы, воздуховоды и кабельные лотки для продолжения");

                XYZ targetPoint = _entity.RawRoutePoints.FirstOrDefault() ?? XYZ.Zero;
                List<MepCurveData> sourceCurves = references
                    .Select(r => uiDoc.Document.GetElement(r))
                    .Where(e => e != null)
                    .Select(e => RouteBuilder.CreateMepCurveData(uiDoc.Document, e, targetPoint))
                    .ToList();

                _entity.SetSourceCurves(sourceCurves);
                _entity.ClearPreview();
                _entity.UserMainStatus = $"Выбрано элементов: {sourceCurves.Count}.";
                _entity.UserHelp = "Теперь укажи точки траектории. Каждая точка сразу добавит новый preview-сегмент.";

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.UserMainStatus = "Не удалось выбрать пучок.";
                _entity.UserHelp = ex.Message;
                return Result.Failed;
            }
            finally
            {
                _window?.Show();
                _window?.Activate();
            }
        }
    }
}
