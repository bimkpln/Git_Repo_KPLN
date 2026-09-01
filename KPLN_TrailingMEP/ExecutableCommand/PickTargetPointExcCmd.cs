using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Common;
using KPLN_TrailingMEP.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_TrailingMEP.ExecutableCommand
{
    internal sealed class PickTargetPointExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;
        private readonly Window _window;

        public PickTargetPointExcCmd(RouteTraceM entity, Window window)
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
                if (!_entity.CanPickPoints)
                {
                    _entity.UserHelp = "Сначала выбери пучок труб, воздуховодов или кабельных лотков.";
                    return Result.Cancelled;
                }

                _window?.Hide();

                while (true)
                {
                    XYZ targetPoint = uiDoc.Selection.PickPoint("Укажи следующую точку траектории. Esc - закончить выбор точек");
                    _entity.AddRawRoutePoint(targetPoint);

                    if (!_entity.HasValidRouteData(out string reason))
                    {
                        _entity.UserHelp = reason;
                        continue;
                    }

                    _entity.BeginInternalRouteChange();
                    try
                    {
                        using (Transaction transaction = new Transaction(uiDoc.Document, "KPLN: добавить сегмент траектории MEP"))
                        {
                            transaction.Start();
                            IReadOnlyList<ElementId> routeIds = RouteBuilder.CreateOrReplacePreviewRoute(
                                uiDoc.Document,
                                _entity.PreviewRouteIds,
                                _entity.GetBundleBasePoint(),
                                _entity.RawRoutePoints,
                                _entity.GetBundleBaseDirection(),
                                _entity.AutoCorrectRoute,
                                _entity.GetAllowedAngles());
                            _entity.SetPreviewRouteIds(routeIds);
                            transaction.Commit();
                        }
                    }
                    finally
                    {
                        _entity.EndInternalRouteChange();
                    }

                    _entity.UserMainStatus = $"Точек траектории: {_entity.RawRoutePoints.Count}.";
                    _entity.UserHelp = "Можно продолжить указывать точки или закончить Esc. После ручной правки линий выбери траекторию заново.";
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return _entity.RawRoutePoints.Count > 0 ? Result.Succeeded : Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.UserMainStatus = "Не удалось указать точку назначения.";
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
