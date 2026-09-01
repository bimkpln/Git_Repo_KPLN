using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Common;
using KPLN_TrailingMEP.Forms.Entities;
using KPLN_Loader.Common;
using System;

namespace KPLN_TrailingMEP.ExecutableCommand
{
    internal sealed class DeletePreviewRouteExcCmd : IExecutableCommand
    {
        private readonly RouteTraceM _entity;

        public DeletePreviewRouteExcCmd(RouteTraceM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            try
            {
                _entity.BeginInternalRouteChange();
                try
                {
                    using (Transaction transaction = new Transaction(doc, "KPLN: удалить линию траектории MEP"))
                    {
                        transaction.Start();
                        RouteBuilder.DeletePreviewRoutes(doc, _entity.PreviewRouteIds);
                        transaction.Commit();
                    }
                }
                finally
                {
                    _entity.EndInternalRouteChange();
                }

                _entity.ClearPreview();
                _entity.UserMainStatus = "Траектория удалена.";
                _entity.UserHelp = string.Empty;
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                _entity.UserMainStatus = "Не удалось удалить линию траектории.";
                _entity.UserHelp = ex.Message;
                return Result.Failed;
            }
        }
    }
}
