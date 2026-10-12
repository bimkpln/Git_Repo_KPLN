using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Forms.Entities;
using System;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class OVVK_SystemManagerExcCmd : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly Action<UIApplication> _action;

        internal OVVK_SystemManagerExcCmd(OVVK_SystemManagerM model, Action<UIApplication> action)
        {
            _model = model;
            _action = action;
        }


        public Result Execute(UIApplication app)
        {
            try
            {
                if (_model.IsClosed || !_model.CurrentDoc.IsValidObject)
                    return Result.Cancelled;

                _model.IsDocumentActive = app.ActiveUIDocument != null
                    && _model.CurrentDoc.Equals(app.ActiveUIDocument.Document);
                if (!_model.IsDocumentActive)
                    return Result.Cancelled;


                _model.SetOperationMessage(string.Empty);
                _action(app);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                _model.SetOperationMessage("Операция не выполнена: " + ex.Message);
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
            finally
            {
                _model.IsBusy = false;
            }
        }
    }
}
