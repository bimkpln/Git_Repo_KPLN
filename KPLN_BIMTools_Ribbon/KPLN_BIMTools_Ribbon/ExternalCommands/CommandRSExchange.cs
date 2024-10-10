using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using RevitServerAPILib;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    /// <summary>
    /// Плагин по импорту моделей с Revit-Server
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandRSExchange : ExchangeService, IExternalCommand, IExecuteByUIApp
    {
        public CommandRSExchange()
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application, RevitDocExchangeEnum.RevitServer);
        }

        public Result ExecuteByUIApp(UIApplication uiapp, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            try
            {
                StartService(uiapp, revitDocExchangeEnum);
            }
            catch (Exception ex)
            {
                PrintError(ex);

                return Result.Failed;
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод открытия и копирования файла по новому пути
        /// </summary>
        private protected override string ExchangeFile(Application app, ModelPath modelPathFrom, DBConfigEntity configEntity, string rsn = "")
        {
            //Апкастинг в настройку для экспорта в RS
            if (configEntity is DBRSConfigData rsConfigData)
            {
                // Подготовка к открытию
                SetOpenOptions(WorksetConfigurationOption.CloseAllWorksets);
                SetSaveAsOptions();

                // Открываем документ по указанному пути
                Document doc = app.OpenDocumentFile(
                    modelPathFrom,
                    _openOptions);

                if (doc != null)
                {
                    string newPath = $"{rsn}{rsConfigData.PathTo}\\{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}.rvt";
                    ModelPath newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath);

                    doc.SaveAs(newModelPath, _saveAsOptions);
                    CurrentDocName = doc.Title;
                    WorksharingUtils.RelinquishOwnership(doc, new RelinquishOptions(true), new TransactWithCentralOptions());
                    doc.Close(false);

                    return newPath;
                }
                else
                    Logger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). Нужно вмешаться человеку");
            }
            else
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBRSConfigData)}");


            return null;
        }
    }
}
