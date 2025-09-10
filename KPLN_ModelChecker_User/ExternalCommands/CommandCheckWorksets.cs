using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckWorksets : AbstrCommand, IExternalCommand
    {
        public CommandCheckWorksets() : base() { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            
            CommandCheck = new CheckWorksets();
            ElemsToCheck = CommandCheck.GetElemsToCheck(doc);

            // Блокирую проверку части РН ПРИ РУЧНОМ ЗАПУСКЕ
            // Для авт. запуска блок не нужен, достаточно проверять часть.
            // Даже если это перезапуск - изначально всё равно нужно вручную запустить и открыть РН
            if (!doc.IsWorkshared)
            {
                TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                {
                    MainContent = $"Проверка рабочих наборов может выполняться ТОЛЬКО с моделями для совместной работы",
                };
                taskDialog.Show();

                return Result.Cancelled;
            }

            Workset[] worksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToArray();
            foreach (Workset ws in worksets)
            {
                if (!ws.IsOpen)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = $"Необходимо загрузить ВСЕ рабочие наборы",
                    };
                    taskDialog.Show();

                    return Result.Cancelled;
                }
            }

            if (ExecuteByUIApp<CheckWorksets>(commandData.Application, true, true, true, true))
                return Result.Succeeded;

            return Result.Cancelled;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}
