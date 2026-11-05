using System;
using System.Windows.Input;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker.Core;
using KPLN_ModelChecker_User.Checks.KR;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.Forms;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.ReadOnly)]
    public sealed class CommandCheckKRIfcMarks : KRCheckCommand, IExternalCommand
    {
        public const string PluginName = "КР: Проверка марок IFC";
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
            => Run<CheckIfcRebarMarks>(data.Application, PluginName, ref message);
    }

    [Transaction(TransactionMode.ReadOnly)]
    public sealed class CommandCheckKRConcreteClass : KRCheckCommand, IExternalCommand
    {
        public const string PluginName = "КР: Проверка класса бетона арматуры";
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
            => Run<CheckRebarConcreteClass>(data.Application, PluginName, ref message);
    }

    public abstract class KRCheckCommand : AbstrCommand
    {
        private string _reportTitle;
        private Cursor _previousCursor;
        protected override bool UseDefaultErrorReporting => false;

        protected Result Run<T>(UIApplication app, string title, ref string message) where T : AbstrKRCheck, new()
        {
            if (!KRCheckAccess.IsAllowed())
            {
                TaskDialog.Show(title, "Проверка доступна только отделам КР и BIM. Проверьте данные пользователя в БД KPLN.");
                return Result.Cancelled;
            }
            var doc = app.ActiveUIDocument?.Document;
            if (doc == null || doc.IsFamilyDocument)
            {
                TaskDialog.Show(title, "Откройте проект RVT.");
                return Result.Cancelled;
            }
            _reportTitle = title;
            _previousCursor = Mouse.OverrideCursor;
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                CommandCheck = null;
                ElemsToCheck = null;
                ExecuteByUIApp<T>(app, onlyErrorType: false, setPluginActivity: false,
                    showMainForm: true, setLastRun: false, showSuccsessText: false);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                TaskDialog.Show(title, "Проверка не завершена.\n" + ex.Message);
                return Result.Failed;
            }
            finally { Mouse.OverrideCursor = _previousCursor; }
        }

        private protected override void PresentResult<T>(UIApplication uiapp, IJsonSerializable config,
            CheckResultStatus checkResultStatus, bool showMainForm, bool setLastRun, bool showSuccsessText)
        {
            // Keep the agreed report window. No ExtensibleStorage/log writes, success HTML or marking.
            var check = (AbstrKRCheck)CommandCheck;
            var report = KRCheckReport.FromData(check.Report.Data);
            Mouse.OverrideCursor = _previousCursor;
            if (showMainForm)
                new KRCheckReportWindow(uiapp, uiapp.ActiveUIDocument.Document,
                    _reportTitle ?? check.PluginName, report).Show();
        }
    }
}
