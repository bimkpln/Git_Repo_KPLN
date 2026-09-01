using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Common;
using KPLN_TrailingMEP.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;

namespace KPLN_TrailingMEP.ExecutableCommand
{
    internal sealed class CreatePreviewRouteExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;

        public CreatePreviewRouteExcCmd(RouteTraceM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            try
            {
                if (!_entity.HasValidRouteData(out string reason))
                {
                    _entity.UserHelp = reason;
                    return Result.Cancelled;
                }

                Document doc = app.ActiveUIDocument.Document;
                XYZ startPoint = _entity.GetBundleBasePoint();

                _entity.BeginInternalRouteChange();
                try
                {
                    using (Transaction transaction = new Transaction(doc, "KPLN: линия детализации траектории MEP"))
                    {
                        transaction.Start();
                        IReadOnlyList<ElementId> routeIds = RouteBuilder.CreateOrReplacePreviewRoute(
                            doc,
                            _entity.PreviewRouteIds,
                            startPoint,
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

                _entity.UserMainStatus = "Траектория создана.";
                _entity.UserHelp = $"Если поправишь линии детализации вручную, выбери траекторию со стилем {RouteBuilder.PreviewLineStyleName} заново перед построением.";
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.UserMainStatus = "Не удалось создать линию детализации траектории.";
                _entity.UserHelp = ex.Message;
                return Result.Failed;
            }
        }
    }
}
