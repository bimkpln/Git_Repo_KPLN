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
    internal sealed class PickPreviewRouteExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;
        private readonly Window _window;

        public PickPreviewRouteExcCmd(RouteTraceM entity, Window window)
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
                if (!_entity.CanPickPreviewRoute)
                {
                    _entity.UserHelp = "Сначала выбери пучок, чтобы плагин понял, от какого торца читать траекторию.";
                    return Result.Cancelled;
                }

                _window?.Hide();
                IList<Reference> references = uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    new RouteCurveSelectionFilter(RouteBuilder.PreviewLineStyleName),
                    $"Выбери линии траектории со стилем {RouteBuilder.PreviewLineStyleName}");

                IReadOnlyList<ElementId> routeIds = references
                    .Select(r => r.ElementId)
                    .Where(id => id != null && !id.Equals(ElementId.InvalidElementId))
                    .ToList();

                _entity.SetPreviewRouteIds(routeIds);
                IReadOnlyList<XYZ> routePoints = RouteBuilder.GetPreviewRoutePoints(uiDoc.Document, routeIds, _entity.GetBundleBasePoint());
                _entity.SetRawRoutePoints(routePoints.Skip(1));

                _entity.UserMainStatus = $"Выбрано сегментов траектории: {routeIds.Count}.";
                _entity.UserHelp = "Траектория принята. После ручной правки линий выбери траекторию заново перед построением.";
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.UserMainStatus = "Не удалось выбрать траекторию.";
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
