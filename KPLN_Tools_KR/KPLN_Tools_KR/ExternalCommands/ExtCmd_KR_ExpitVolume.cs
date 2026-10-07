using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_KR.Forms;
using System;

namespace KPLN_Tools_KR.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_KR_ExpitVolume : IExternalCommand
    {
        internal const string PluginName = "Получение объема котлована";
        private static ExpitVolume _activeWindow;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                TaskDialog.Show(
                    PluginName,
                    "Команда доступна в Revit 2024 и новее, поскольку работает с топотелами.");
                return Result.Cancelled;
#else
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc?.Document;

                if (doc == null)
                {
                    message = "Не удалось получить активный документ Revit.";
                    return Result.Failed;
                }

                if (doc.IsFamilyDocument)
                {
                    TaskDialog.Show(PluginName, "Команда работает только в документе проекта.");
                    return Result.Cancelled;
                }

                if (_activeWindow != null && _activeWindow.IsLoaded)
                {
                    _activeWindow.Activate();
                    return Result.Succeeded;
                }

                _activeWindow = new ExpitVolume(uidoc);
                _activeWindow.Closed += (sender, args) => _activeWindow = null;
                _activeWindow.Show();
                return Result.Succeeded;
#endif
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(PluginName, "Ошибка:\n" + ex.Message);
                return Result.Failed;
            }
        }
    }
}