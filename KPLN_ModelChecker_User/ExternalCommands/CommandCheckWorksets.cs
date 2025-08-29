using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.WPFItems;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckWorksets : AbstrCommand<CommandCheckWorksets>, IExternalCommand
    {
        internal const string PluginName = "Проверка рабочих наборов";

        private CheckWorksets _checkWorksets;
        private Element[] _elemsToCheck;

        public CommandCheckWorksets() : base() { }

        internal CommandCheckWorksets(ExtensibleStorageEntity esEntity) : base(esEntity) { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _checkWorksets = new CheckWorksets(commandData.Application);
            _elemsToCheck = _checkWorksets.GetElemsToCheck();

            Document doc = commandData.Application.ActiveUIDocument.Document;

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

            return ExecuteByUIApp(commandData.Application, true, true, true);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp, bool setPluginActivity, bool showMainForm, bool showSuccsessText)
        {
            if (setPluginActivity)
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            if (_checkWorksets == null || _elemsToCheck == null)
            {
                _checkWorksets = new CheckWorksets(uiapp);
                _elemsToCheck = _checkWorksets.GetElemsToCheck();
            }

            CheckerEntities = _checkWorksets.ExecuteCheck(_elemsToCheck);
            if (CheckerEntities != null && CheckerEntities.Length > 0 && showMainForm)
                ReportCreatorAndDemonstrator(uiapp);
            else if (showSuccsessText)
                HtmlOutput.Print($"[{ESEntity.CheckName}] Предупреждений не найдено :)", MessageType.Success);

            return Result.Succeeded;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}
