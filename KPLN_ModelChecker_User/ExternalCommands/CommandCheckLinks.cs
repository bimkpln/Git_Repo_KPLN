using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckLinks : AbstrCommand<CommandCheckLinks>, IExternalCommand
    {
        internal const string PluginName = "Проверка связей";

        private CheckLinks _checkLinks;
        private Element[] _elemsToCheck;

        public CommandCheckLinks() : base() {}

        internal CommandCheckLinks(ExtensibleStorageEntity esEntity) : base(esEntity) {}

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _checkLinks = new CheckLinks(commandData.Application);
            _elemsToCheck = _checkLinks.GetElemsToCheck();

            // Блокирую проверку части линков ПРИ РУЧНОМ ЗАПУСКЕ
            // Для авт. запуска блок не нужен, достаточно проверять часть.
            // Даже если это перезапуск - изначально всё равно нужно вручную запустить и подгрузить линки
            foreach (RevitLinkInstance rli in _elemsToCheck.Cast<RevitLinkInstance>())
            {
                Document linkDoc = rli.GetLinkDocument();
                if (linkDoc == null)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = $"Необходимо загрузить ВСЕ связи. Проверь диспетчер Revit-связей",
                    };
                    taskDialog.Show();

                    return Result.Cancelled;
                }
            }

            return ExecuteByUIApp(commandData.Application, true, true, true, true);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp, bool setPluginActivity, bool showMainForm, bool setLastRun, bool showSuccsessText)
        {
            if (setPluginActivity)
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            if (_checkLinks == null || _elemsToCheck == null)
            {
                _checkLinks = new CheckLinks(uiapp);
                _elemsToCheck = _checkLinks.GetElemsToCheck();
            }
           
            CheckerEntities = _checkLinks.ExecuteCheck(_elemsToCheck);
            if (CheckerEntities != null && CheckerEntities.Length > 0 && showMainForm)
                ReportCreatorAndDemonstrator(uiapp, setLastRun);
            else if (showSuccsessText)
            {
                // Логируем последний запуск (отдельно, если все было ОК, а потом всплыли ошибки)
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESEntity.ESBuilderRun, DateTime.Now));
                
                HtmlOutput.Print($"[{ESEntity.CheckName}] Предупреждений не найдено :)", MessageType.Success);
            }

            return Result.Succeeded;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}