using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_DBWorker.Core;
using System;
using System.Threading;
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
            return ExecuteByUIApp(commandData.Application, RevitDocExchangeEnum.Revit);
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
            if (!(configEntity is DBRVTConfigData rsConfigData))
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBRVTConfigData)}");


            // Подготовка к открытию
            SetOpenOptions(WorksetConfigurationOption.CloseAllWorksets);
            SetSaveAsOptions(rsConfigData);

            
            // Открываем документ по указанному пути
            Document doc = null;
            try
            {
                try
                {
                    // Добавил задержку, т.к. бывает файл не хочет открыться, и ошибка "was thrown by Revit or by one of its external applications"
                    Thread.Sleep(2000);

                    doc = app.OpenDocumentFile(
                        modelPathFrom,
                        _openOptions);
                }
                catch (Autodesk.Revit.Exceptions.FileNotFoundException)
                {
                    string modelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom);
                    string msg = $"Путь к файлу {modelPath} - не существует. Внимательно проверь путь и наличие модели по указанному пути";
                    Print(msg, MessageType.Warning);
                    Module.CurrentLogger.Error(msg);

                    return null;
                }
                catch (Exception ex)
                {
                    Module.CurrentLogger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). Нужно вмешаться человеку, " +
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

                return mutablePath;
            }
            catch (Exception ex)
            {
                Module.CurrentLogger.Error(
                    $"Не выгрузить RVT ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). " +
                    $"Ошибка: {ex.Message}");

                return null;
            }
            finally
            {
                if (doc != null && doc.IsValidObject)
                {
                    try
                    {
                        doc.Close(false);
                    }
                    catch (Exception ex)
                    {
                        Module.CurrentLogger.Error($"Не удалось закрыть документ после RVT-обмена. Ошибка: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Замена пути в соответсвии с требованиями конфига
        /// </summary>
        private string NameMutabledByConfig(string newPath, string docTitle, DBRVTConfigData rsConfigData)
        {
            if (string.IsNullOrEmpty(rsConfigData.NameChangeFind))
                return newPath;

            string configNameChangeFind = rsConfigData.NameChangeFind;
            // Для обработки старых конфигов, чтобы не писал об ошибках
            if (configNameChangeFind == "🔐")
                return newPath;

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
