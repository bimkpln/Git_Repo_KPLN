using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_BIMTools_Ribbon.Forms;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_OpenDocHandler;
using KPLN_Library_OpenDocHandler.Core;
using KPLN_Library_PluginActivityWorker;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Общий сервис по подготовке к работе и работе с файлами для обмена
    /// </summary>
    public class ExchangeService : IFieldChangedNotifier
    {
        public event EventHandler<FieldChangedEventArgs> FieldChanged;

        private protected string _sourceProjectName;
        private protected OpenOptions _openOptions;
        private protected SaveAsOptions _saveAsOptions;

        private static RevitDocExchangesDbService _revitDocExchangesDbService;

        private string _currentDocName;

        internal static RevitDocExchangesDbService RevitDocExchangesDbService
        {
            get
            {
                if (_revitDocExchangesDbService == null)
                    _revitDocExchangesDbService = (RevitDocExchangesDbService)new CreatorRevitDocExchangesDbService().CreateService();
                return _revitDocExchangesDbService;
            }
        }

        internal static UIControlledApplication RevitUIControlledApp { get; set; }

        /// <summary>
        /// Метка сервиса о том, что он запускается автоматически
        /// </summary>
        internal static bool IsAutoStart { get; set; } = false;

        /// <summary>
        /// Счтетчик успешно отработанных процессов
        /// </summary>
        internal int CountProcessedDocs { get; set; } = 0;

        /// <summary>
        /// Счтетчик файлов для обработки
        /// </summary>
        internal int CountSourceDocs { get; set; } = 0;

        /// <summary>
        /// Имя документа в текущем процессе обработки
        /// </summary>
        internal string CurrentDocName
        {
            get => _currentDocName;
            set
            {
                if (_currentDocName != value)
                {
                    _currentDocName = value;
                    OnFieldChanged(new FieldChangedEventArgs(value));
                }
            }
        }

        /// <summary>
        /// Установка общих параметров для запуска
        /// </summary>
        internal static void SetStaticEnvironment(UIControlledApplication application)
        {
            RevitUIControlledApp = application;
        }

        /// <summary>
        /// Старт сервиса по обмену моделями
        /// </summary>
        private protected void StartService(UIApplication uiapp, RevitDocExchangeEnum revitDocExchangeEnum, string pluginName)
        {
            string configNames = "Не определено";
            string docExchangeModuleName = "Не определено";

            // Подготовка коллекции для экспорта
            DBRevitDocExchangesWrapper[] dbRevitDocExchanges = null;
            if (IsAutoStart)
            {
                docExchangeModuleName = $"Автостарт: {revitDocExchangeEnum}";

                IEnumerable<int> docExchIdsFromModuleAS = DBMainService.ModuleAutostartDbService
                    .GetDBModuleAutostarts_ByUserAndRVersionAndTable(DBMainService.CurrentDBUser.Id, ModuleData.RevitVersion, DB_Enumerator.RevitDocExchanges.ToString())
                    .Select(mas => mas.DBTableKeyId);
                if (docExchIdsFromModuleAS.Count() == 0)
                    return;

                IEnumerable<DBRevitDocExchanges> docExcs = RevitDocExchangesDbService
                    .GetDBRevitDocExchanges_ByIdCol(docExchIdsFromModuleAS);
                if (docExcs.Count() == 0)
                    return;
                
                dbRevitDocExchanges = docExcs
                    .Select(dExc => new DBRevitDocExchangesWrapper(dExc))
                    .ToArray();

                int prjId = docExcs.FirstOrDefault().ProjectId;
                DBProject dBProject = DBMainService.ProjectDbService.GetDBProject_ByProjectId(prjId);
                _sourceProjectName = dBProject.Name;

                configNames = $"Автостарт: {string.Join("; ", dbRevitDocExchanges.Select(de => de.SettingName))}";

                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"Автостарт: {pluginName}", ModuleData.ModuleName).ConfigureAwait(false);
            }
            else
            {
                docExchangeModuleName = $"{revitDocExchangeEnum}";

                ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm(null, ModuleData.RevitVersion, true);
                if (!(bool)selectedProjectForm.ShowDialog())
                    return;


                DBProject dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
                _sourceProjectName = dBProject.Name;

                ConfigDispatcher configDispatcher = new ConfigDispatcher(dBProject, revitDocExchangeEnum, false);
                if (!(bool)configDispatcher.ShowDialog())
                    return;

                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(pluginName, ModuleData.ModuleName).ConfigureAwait(false);

                configNames = string.Join("; ", configDispatcher.SelectedDBExchWrappers.Select(ent => ent.SettingName));
                dbRevitDocExchanges = configDispatcher.SelectedDBExchWrappers;
            }

            if (dbRevitDocExchanges == null)
                return;

            using (UIContrAppSubscriber subscriber = new UIContrAppSubscriber(RevitUIControlledApp, Module.CurrentLogger, this))
            {
                // Локальный try, чтобы гарантированно отписаться от событий. Cath - кидает ошибку выше
                try
                {
                    Module.CurrentLogger.Info($"Старт экспорта: [{docExchangeModuleName}].\nКонфигурация/-ии: [{configNames}]");

                    foreach (DBRevitDocExchangesWrapper currentDocExchEnt in dbRevitDocExchanges)
                    {
                        SQLiteService sqliteService = new SQLiteService(currentDocExchEnt.SettingDBFilePath, revitDocExchangeEnum);
                        IEnumerable<DBConfigEntity> configs = sqliteService.GetConfigItems();
                        foreach (DBConfigEntity config in configs)
                        {
                            List<string> fileFromPathes = PreparePathesToOpen(config.PathFrom);
                            if (fileFromPathes != null)
                            {
                                CountSourceDocs += fileFromPathes.Count;
                                foreach (string fileFromPath in fileFromPathes)
                                {
                                    string newFilePath = string.Empty;
                                    ModelPath docFromModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(fileFromPath);

                                    // Проверяю КУДА копирвать.
                                    // Это папка, если нет - то ревит-сервер
                                    bool isRevitServerFile = false;
                                    bool isKPLNServerFile = false;
                                    if (Directory.Exists(config.PathTo))
                                        isKPLNServerFile = true;
                                    // Убеждаюсь и обрабатываю ревит-сервер
                                    else if (CheckPathFoRevitServer(config.PathTo))
                                        isRevitServerFile = true;
                                    
                                    
                                    // Если ничего из вышеописанного - то ошибка
                                    if (isRevitServerFile == isKPLNServerFile)
                                    {
                                        Module.CurrentLogger.Error($"Файл {config.PathFrom} не удалось определить путь для сохранения {config.PathTo}.\n");
                                        continue;
                                    }

                                    // Часто встречаются фантомные ошибки открытия, особенно с RS. Ввожу итерации
                                    int exchIteration = 1;
                                    int maxExchIteration = 2;
                                    while (exchIteration <= maxExchIteration)
                                    {
                                        // Запускаю экспорт
                                        if (isKPLNServerFile)
                                            newFilePath = ExchangeFile(uiapp.Application, docFromModelPath, config);
                                        else if (isRevitServerFile)
                                            newFilePath = ExchangeFile(uiapp.Application, docFromModelPath, config, "RSN:");

                                        // Проверка результатов итерации
                                        if (newFilePath != null && !string.IsNullOrEmpty(newFilePath))
                                        {
                                            CountProcessedDocs++;
                                            break;
                                        }

                                        Module.CurrentLogger.Debug($"Файл {config.Name} не экспортирован. Ошибки описаны выше. Выполнена итерация {exchIteration} из {maxExchIteration} возможных");
                                        exchIteration++;
                                    }


                                    // След. итерации не помогли, выхожу
                                    if (newFilePath == null || string.IsNullOrEmpty(newFilePath))
                                        Module.CurrentLogger.Error($"Файл {config.Name} не экспортирован (количество попыток - {maxExchIteration}). Ошибки описаны выше.\n");
                                }
                            }
                            // Все равно добавляю 1, чтобы попало в отчет
                            else CountSourceDocs++;
                        }
                    }

                    SendResultMsg($"Плагин экспорта [{docExchangeModuleName}]", configNames);
                    Module.CurrentLogger.Info($"Работа плагина [{docExchangeModuleName}] завершена.\n");
                }
                catch (Exception ex)
                {
                    SendResultMsg($"Плагин экспорта [{docExchangeModuleName}]", configNames);
                    Module.CurrentLogger.Error($"Работа плагина [{docExchangeModuleName}] ЭКСТРЕННО завершена. Ошибка: {ex.Message}\n");
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Метод обмена файлами
        /// </summary>
        private protected virtual string ExchangeFile(Application app, ModelPath modelPathFrom, DBConfigEntity configEntity, string rsn = "")
        {
            throw new NotImplementedException("Ошибка реализации структуры! Нужно переопределить метод ExchangeFiles для каджого экспортера");
        }

        /// <summary>
        /// Подготовка опций к открытию
        /// </summary>
        private protected void SetOpenOptions(WorksetConfigurationOption worksetConfigurationOption)
        {
            _openOptions = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets };
            _openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(worksetConfigurationOption));
        }

        /// <summary>
        /// Подготовка опций к открытию с указанием рабочих наборов
        /// </summary>
        private protected void SetOpenOptions(IList<WorksetId> worksetIds)
        {
            _openOptions = new OpenOptions()
            {
                DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            };
            WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openConfig.Open(worksetIds);
            _openOptions.SetOpenWorksetsConfiguration(openConfig);
        }

        /// <summary>
        /// Подготовка опций к сохранению
        /// </summary>
        private protected void SetSaveAsOptions(DBRVTConfigData dBRSConfigData)
        {
            int backupTempForOldPlugin;
            if (dBRSConfigData.MaxBackup == -1)
                backupTempForOldPlugin = 10;
            else
                backupTempForOldPlugin = dBRSConfigData.MaxBackup;

            _saveAsOptions = new SaveAsOptions()
            {
                OverwriteExistingFile = true,
                MaximumBackups = backupTempForOldPlugin,
            };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions()
            {
                SaveAsCentral = true,
                OpenWorksetsDefault = SimpleWorksetConfiguration.AskUserToSpecify
            };
            _saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);
        }


        /// <summary>
        /// Обработчик события
        /// </summary>
        /// <param name="e"></param>
        private void OnFieldChanged(FieldChangedEventArgs e) =>
            FieldChanged?.Invoke(this, e);

        /// <summary>
        /// Отправка результата пользователю в месенджер
        /// </summary>
        private void SendResultMsg(string moduleName, string configNames)
        {
            if (CountProcessedDocs < CountSourceDocs || CountProcessedDocs == 0)
            {
                BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                    DBMainService.CurrentDBUser,
                    $"Модуль: [b]{moduleName}\n[/b]" +
                    $"Анализируемые конфигурации: {configNames}\n" +
                    $"Статус: Отработано с ошибками.\n" +
                    $"Метрик производительности: Выгружено {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта: [b]{_sourceProjectName}[/b]\n" +
                    $"Ошибки: См. файл логов у пользователя {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name}.\n" +
                    $"Путь к логам у пользователя: C:\\KPLN_Temp\\KPLN_Logs\\{ModuleData.RevitVersion}");
            }
            else
            {
                BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                    DBMainService.CurrentDBUser,
                    $"Модуль: [b]{moduleName}\n[/b]" +
                    $"Анализируемые конфигурации: {configNames}\n" +
                    $"Статус: Отработано без ошибок.\n" +
                    $"Метрик производительности: Обработано {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта: [b]{_sourceProjectName}[/b]");
            }
        }

        /// <summary>
        /// Подготовка путей к открытию в Revit
        /// </summary>
        /// <param name="pathFrom"></param>
        /// <returns></returns>
        private List<string> PreparePathesToOpen(string pathFrom)
        {
            List<string> fileFromPathes = new List<string>();
            string[] pathParts = pathFrom.Split('\\');

            // Проверяю, что это файл, если нет - то нужно забрать ВСЕ файлы из папки
            if (System.IO.File.Exists(pathFrom))
                fileFromPathes.Add(pathFrom);
            #region Проверяю, что это папка, если нет - то нужно забрать ВСЕ файлы из ревит-сервера (ПО СТАРОМУ ВАРИАНТУ - КУЧА КОНФИГОВ, УДАЛИТЬ ПОЗЖЕ!!!)
            else if (Directory.Exists(pathFrom))
                fileFromPathes = Directory.GetFiles(pathFrom, "*" + ".rvt").ToList<string>();
            #endregion
            #region Обработка Revit-Server, чтобы забрать файл или ВСЕ файлы из папки (ПО СТАРОМУ ВАРИАНТУ - КУЧА КОНФИГОВ, УДАЛИТЬ ПОЗЖЕ!!!)
            // https://www.nuget.org/packages/RevitServerAPILib
            else if (string.IsNullOrEmpty(pathParts[0]))
            {
                string rsHostName = pathParts[2];
                int pathPartsLenght = pathParts.Length;
                if (rsHostName == null)
                {
                    Module.CurrentLogger.Error($"Ошибка заполнения пути для копирования с Revit-Server: ({pathFrom}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'");
                    return null;
                }
                try
                {
                    RevitServer server = new RevitServer(rsHostName, ModuleData.RevitVersion);
                    // Проверяю ссылку на конечный файл. Добавляю файл
                    if (pathFrom.ToLower().Contains("rvt"))
                    {
                        fileFromPathes.Add($"RSN:{pathFrom}");
                    }
                    // Значит ссылка на папку. Добавляю файлы
                    else
                    {
                        FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 3));
                        foreach (var model in folderContents.Models)
                        {
                            fileFromPathes.Add($"RSN:{pathFrom}{model.Name}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Module.CurrentLogger.Error($"Ошибка открытия Revit-Server ({pathFrom}):\n{ex.Message}");
                    return null;
                }
            }
            #endregion
            // Обработка моделей с Revit-Server (все пути уже заранее готовы)
            else if (pathFrom.StartsWith("RSN:"))
                fileFromPathes.Add($"{pathFrom}");

            if (fileFromPathes.Count == 0)
            {
                Module.CurrentLogger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return null;
            }

            return fileFromPathes;
        }

        /// <summary>
        /// Метод проверки пути на РС на наличие папки (RevitServerAPILib сам её может создать, но мне это не подходит)
        /// </summary>
        /// <param name="pathTo">Путь для проверки</param>
        /// <returns></returns>
        private bool CheckPathFoRevitServer(string pathTo)
        {
            if (Directory.Exists(pathTo))
                return false;

            string[] pathParts = pathTo.Split('\\');
            string rsHostName = pathParts[2];
            int pathPartsLenght = pathParts.Length;
            RevitServer server = new RevitServer(rsHostName, ModuleData.RevitVersion);
            if (server != null)
            {
                try
                {
                    FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 3));
                    if (folderContents != null)
                        return true;
                    else
                        return false;

                }
                catch (System.Net.WebException wex)
                {
                    if (wex.Message.Contains("404"))
                        return false;
                    else
                        throw wex;
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }
            else
                return false;
        }
    }
}
