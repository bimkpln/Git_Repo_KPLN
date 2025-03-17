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
    internal class CommandRVTExchange : ExchangeService, IExternalCommand, IExecuteByUIApp
    {
        internal const string PluginName = "RVT: Обмен";

        public CommandRVTExchange()
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
                StartService(uiapp, revitDocExchangeEnum, PluginName);
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
            if (configEntity is DBRVTConfigData rsConfigData)
            {
                // Подготовка к открытию
                SetOpenOptions(WorksetConfigurationOption.CloseAllWorksets);
                SetSaveAsOptions(rsConfigData);

                Document doc = null;
                // Открываем документ по указанному пути
                try
                {
                    doc = app.OpenDocumentFile(
                        modelPathFrom,
                        _openOptions);
                }
                catch (Autodesk.Revit.Exceptions.FileNotFoundException)
                {
                    string modelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom);
                    string msg = $"Путь к файлу {modelPath} - не существует. Внимательно проверь путь и наличие модели по указанному пути";
                    Print(msg, MessageType.Warning);
                    Logger.Error(msg);
                    
                    return null;
                }
                catch (Exception ex)
                {
                    Logger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). Нужно вмешаться человеку, " +
                        $"ошибка при открытии: {ex.Message}");
                    
                    return null;
                }

                string docTitle = doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0];
                string newPath = $"{rsn}{rsConfigData.PathTo}\\{docTitle}.rvt";
                string mutablePath = NameMutabledByConfig(newPath, docTitle, rsConfigData);

                ModelPath newMutableModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(mutablePath);

                doc.SaveAs(newMutableModelPath, _saveAsOptions);
                CurrentDocName = doc.Title;
                WorksharingUtils.RelinquishOwnership(doc, new RelinquishOptions(true), new TransactWithCentralOptions());
                doc.Close(false);

                return mutablePath;
            }
            else
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBRVTConfigData)}");
        }

        /// <summary>
        /// Замена пути в соответсвии с требованиями конфига
        /// </summary>
        private string NameMutabledByConfig(string newPath, string docTitle, DBRVTConfigData rsConfigData)
        {
            if (string.IsNullOrEmpty(rsConfigData.NameChangeFind))
                return newPath;

            string configNameChangeFind = rsConfigData.NameChangeFind;
            if (!newPath.Contains(configNameChangeFind))
            {
                Print(
                    $"Внимание - в модели по пути \'{docTitle}\' нет совпадения в имени \'{configNameChangeFind}\' для замены на \'{rsConfigData.NameChangeSet}\'. " +
                        $"Модель сохранена со СТАРЫМ именем.", 
                    MessageType.Warning);
                return newPath;
            }
            else
            {
                string mutablePath = newPath.Replace(configNameChangeFind, rsConfigData.NameChangeSet);
                return mutablePath;
            }
        }
    }
}
