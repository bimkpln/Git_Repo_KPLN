using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Common;
using KPLN_TrailingMEP.ExternalCommands;
using KPLN_TrailingMEP.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TrailingMEP.ExecutableCommand
{
    internal sealed class BuildRouteExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;

        public BuildRouteExcCmd(RouteTraceM entity)
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
                if (_entity.IsPreviewRouteChanged)
                {
                    _entity.UserMainStatus = "Траектория изменилась.";
                    _entity.UserHelp = "Выбери траекторию заново, чтобы принять новую геометрию перед построением.";
                    return Result.Cancelled;
                }

                XYZ basePoint = _entity.GetBundleBasePoint();
                IReadOnlyList<XYZ> routePoints = RouteBuilder.GetPreviewRoutePoints(doc, _entity.PreviewRouteIds, basePoint);
                if (routePoints.Count < 2)
                {
                    _entity.UserHelp = "Линии траектории не найдены. Создай траекторию заново.";
                    _entity.ClearPreview();
                    return Result.Cancelled;
                }

                IReadOnlyList<ElementId> createdIds;
                _entity.BeginInternalRouteChange();
                try
                {
                    using (Transaction transaction = new Transaction(doc, "KPLN: дотянуть MEP трассу"))
                    {
                        transaction.Start();
                        createdIds = RouteBuilder.BuildExtensions(doc, _entity.SourceCurves, routePoints);

                        if (_entity.DeletePreviewAfterBuild)
                        {
                            RouteBuilder.DeletePreviewRoutes(doc, _entity.PreviewRouteIds);
                            _entity.ClearPreview();
                        }

                        transaction.Commit();
                    }
                }
                finally
                {
                    _entity.EndInternalRouteChange();
                }

                uiDoc.Selection.SetElementIds(createdIds.ToList());
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(RouteTraceExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                _entity.UserMainStatus = $"Обработано элементов: {createdIds.Count}.";
                _entity.UserHelp = string.Empty;
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.UserMainStatus = "Не удалось построить продолжение трассы.";
                _entity.UserHelp = ex.Message;
                return Result.Failed;
            }
        }
    }
}
