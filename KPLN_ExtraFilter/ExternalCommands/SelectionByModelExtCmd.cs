using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Forms;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using System;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectionByModelExtCmd : IExternalCommand
    {
        /// <summary>
        /// Имя плагина. Использую в KPLN_DefaultPanelExtension_Modify
        /// </summary>
        public const string PluginName = "Дерево элементов";

        /// <summary>
        /// Метод запуска плагина за пределами кнопки в ревит
        /// </summary>
        /// <param name="uiapp">UIApplication для запуска</param>
        /// <param name="viewFilterMode">Предварительная настройка фильтрации элементов</param>
        /// <param name="isUpdated">Настраивать ли обновления в окне</param>
        /// <returns></returns>
        public Result ExecuteByUIApp(UIApplication uiapp, ViewFilterMode viewFilterMode, bool isUpdateble)
        {
            try
            {
                SelectionByModel mainForm = new SelectionByModel(uiapp, viewFilterMode, isUpdateble);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);

                if (isUpdateble)
                    mainForm.Show();
                else
                    mainForm.ShowDialog();

                // Счетчик факта запуска
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) => ExecuteByUIApp(commandData.Application, ViewFilterMode.CurrentView, true);
    }
}
