using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_BIMTools_Ribbon.Forms;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_OpenDocHandler;
using KPLN_Library_OpenDocHandler.Core;
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

        private string _currentDocName;

        internal static UIControlledApplication RevitUIControlledApp { get; set; }

        /// <summary>
        /// Метка сервиса о том, что он запускается автоматически
        /// </summary>
        internal static bool IsAutoStart { get; set; } = false;

        internal static DateTime? AutoStartTimeUtc { get; set; }

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
        /// Поиск части имени в РН
        /// </summary>
        /// <param name="str"></param>
        /// <param name="userQuery"></param>
        /// <returns></returns>
        internal static bool WSName_IsMatchByRules(string rnName, string ruleString)
        {
            if (string.IsNullOrWhiteSpace(rnName) || string.IsNullOrWhiteSpace(ruleString))
                return false;

            string[] rules = ruleString.Split(new[] { '~' }, StringSplitOptions.RemoveEmptyEntries);

            bool includeMatch = false;

            foreach (string raw in rules)
            {
                string rule = raw.Trim();
                if (string.IsNullOrWhiteSpace(rule))
                    continue;

                // !abc! -> НЕ змяшчае
                if (rule.Length >= 2 && rule.StartsWith("!") && rule.EndsWith("!"))
                {
                    string value = rule.Substring(1, rule.Length - 2).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                        continue;

                    if (rnName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
                        return false;
                }
                // *abc* -> змяшчае
                else if (rule.Length >= 2 && rule.StartsWith("*") && rule.EndsWith("*"))
                {
                    string value = rule.Substring(1, rule.Length - 2).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                        continue;

                    if (rnName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
                        includeMatch = true;
                }
                // abc -> пачынаецца з
                else
                {
                    if (rnName.StartsWith(rule, StringComparison.OrdinalIgnoreCase))
                        includeMatch = true;
                }
            }

            return includeMatch;
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
            DBModuleAutostart[] moduleAutostarts = new DBModuleAutostart[0];
            if (IsAutoStart)
            {
                docExchangeModuleName = $"Автостарт: {revitDocExchangeEnum}";

                moduleAutostarts = SQLiteMainService
                    .SQLiteModuleAutostartServiceInst
                    .GetDueDBModuleAutostarts(SQLiteMainService.CurrentDBUser.Id, ModuleData.RevitVersion, 80,
                        DBEnumerator.RevitDocExchanges.ToString(), AutoStartTimeUtc ?? DateTime.UtcNow);
                IEnumerable<int> docExchIdsFromModuleAS = moduleAutostarts.Select(item => item.DBTableKeyId);
                if (docExchIdsFromModuleAS.Count() == 0)
                    return;

                IEnumerable<DBRevitDocExchanges> docExcs = SQLiteMainService
                    .SQLiteRevitDocExchangesServiceInst
                    .GetDBRevitDocExchanges_ByIdCol(docExchIdsFromModuleAS)
                    .Where(item => item.RevitDocExchangeType == revitDocExchangeEnum.ToString())
                    .ToArray();
                if (docExcs.Count() == 0)
                    return;

                dbRevitDocExchanges = docExcs
                    .Select(dExc => new DBRevitDocExchangesWrapper(dExc))
                    .ToArray();

                int prjId = docExcs.FirstOrDefault().ProjectId;
                DBProject dBProject = SQLiteMainService.SQLitePrjServiceInst.GetDBProject_ByProjectId(prjId);
                _sourceProjectName = dBProject.Name;

                configNames = $"Автостарт: {string.Join("; ", dbRevitDocExchanges.Select(de => de.SettingName))}";

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

                configNames = string.Join("; ", configDispatcher.SelectedDBExchWrappers.Select(ent => ent.SettingName));
                dbRevitDocExchanges = configDispatcher.SelectedDBExchWrappers;
            }

            if (dbRevitDocExchanges == null)
                return;

            List<ExchangeNotificationResult> notificationResults = new List<ExchangeNotificationResult>();
            ExchangeNotificationResult currentResult = null;

            using (UIContrAppSubscriber subscriber = new UIContrAppSubscriber(RevitUIControlledApp, Module.CurrentLogger, this))
            {
                // Локальный try, чтобы гарантированно отписаться от событий. Cath - кидает ошибку выше
                try
                {
                    Module.CurrentLogger.Info($"Старт экспорта: [{docExchangeModuleName}].\nКонфигурация/-ии: [{configNames}]");

                    foreach (DBRevitDocExchangesWrapper currentDocExchEnt in dbRevitDocExchanges)
                    {
                        DBModuleAutostart assignment = moduleAutostarts.FirstOrDefault(item => item.DBTableKeyId == currentDocExchEnt.Id
                            && item.ProjectId == currentDocExchEnt.CurrentDBRevitDocExchanges.ProjectId);
                        if (IsAutoStart)
                        {
                            if (assignment == null)
                                continue;

                            // Сначала фиксируем попытку в БД. Без успешной записи выгрузка не начинается.
                            assignment = SQLiteMainService.SQLiteModuleAutostartServiceInst.TryStartExport(assignment, DateTime.UtcNow);
                            if (assignment == null)
                                continue;
                        }

                        currentResult = new ExchangeNotificationResult
                        {
                            Configuration = currentDocExchEnt.CurrentDBRevitDocExchanges,
                            Autostart = assignment,
                            SourceStart = CountSourceDocs,
                            ProcessedStart = CountProcessedDocs,
                            ProjectName = _sourceProjectName,
                        };
                        notificationResults.Add(currentResult);

                        try
                        {
                            // Автозапуск может содержать конфигурации разных проектов.
                            _sourceProjectName = SQLiteMainService.SQLitePrjServiceInst
                                .GetDBProject_ByProjectId(currentResult.Configuration.ProjectId)?.Name ?? "Не определено";
                            currentResult.ProjectName = _sourceProjectName;

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
                                        Module.CurrentLogger.Info($"Приступаю к экспорту файла {ModelPathUtils.ConvertModelPathToUserVisiblePath(docFromModelPath)}");

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
                        catch (Exception ex)
                        {
                            currentResult.Failed = true;
                            Module.CurrentLogger.Error($"Ошибка конфигурации [{currentDocExchEnt.SettingName}]: {ex.Message}");
                            if (!IsAutoStart)
                                throw;
                        }
                        finally
                        {
                            currentResult.SourceCount = CountSourceDocs - currentResult.SourceStart;
                            currentResult.ProcessedCount = CountProcessedDocs - currentResult.ProcessedStart;
                            if (IsAutoStart && assignment != null)
                            {
                                try
                                {
                                    // Завершение фиксируем и после ошибок; расписание учитывает только начало.
                                    if (!SQLiteMainService.SQLiteModuleAutostartServiceInst.CompleteExport(assignment, DateTime.UtcNow))
                                        Module.CurrentLogger.Error($"Не записано завершение [{currentDocExchEnt.SettingName}]: назначение удалено или уже запущена новая попытка.");
                                }
                                catch (Exception ex)
                                {
                                    currentResult.Failed = true;
                                    Module.CurrentLogger.Error($"Не удалось записать завершение [{currentDocExchEnt.SettingName}]: {ex.Message}");
                                }
                            }
                            currentResult = null;
                        }
                    }

                    Module.CurrentLogger.Info($"Работа плагина [{docExchangeModuleName}] завершена.\n");
                }
                catch (Exception ex)
                {
                    if (currentResult != null)
                    {
                        currentResult.SourceCount = CountSourceDocs - currentResult.SourceStart;
                        currentResult.ProcessedCount = CountProcessedDocs - currentResult.ProcessedStart;
                        currentResult.Failed = true;
                    }

                    Module.CurrentLogger.Error($"Работа плагина [{docExchangeModuleName}] ЭКСТРЕННО завершена. Ошибка: {ex.Message}\n");
                    throw;
                }
                finally
                {
                    SendResultMessages($"Плагин экспорта [{docExchangeModuleName}]", notificationResults);
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
        private sealed class ExchangeNotificationResult
        {
            internal DBRevitDocExchanges Configuration { get; set; }
            internal DBModuleAutostart Autostart { get; set; }
            internal string ProjectName { get; set; }
            internal int SourceStart { get; set; }
            internal int ProcessedStart { get; set; }
            internal int SourceCount { get; set; }
            internal int ProcessedCount { get; set; }
            internal bool Failed { get; set; }
        }

        private void SendResultMessages(string moduleName, List<ExchangeNotificationResult> results)
        {
            // Одна сводка по всем конфигам запуска для каждого фактического получателя.
            foreach (NotificationBatch batch in BuildNotificationBatches(results, message => Module.CurrentLogger.Error(message)))
                SendResultMsg(moduleName, batch.Results.ToArray(), batch.Recipient,
                    string.Join("\n", batch.RecipientNotes.Distinct()));
        }

        private sealed class NotificationBatch
        {
            internal DBUser Recipient { get; set; }
            internal List<ExchangeNotificationResult> Results { get; } = new List<ExchangeNotificationResult>();
            internal List<string> RecipientNotes { get; } = new List<string>();
        }

        private static NotificationBatch[] BuildNotificationBatches(IEnumerable<ExchangeNotificationResult> results, Action<string> logError)
        {
            var batches = new Dictionary<int, NotificationBatch>();
            foreach (ExchangeNotificationResult result in results)
            {
                try
                {
                    // Группируем после подстановки отсутствующего пользователя и замены rbim на tkutsko.
                    DBUser recipient = ResolveNotificationRecipient(result.Autostart?.NotificationUserId, out string recipientInfo);
                    if (recipient == null)
                    {
                        logError($"Не удалось отправить уведомление по конфигурации [{result.Configuration.SettingName}]: {recipientInfo}");
                        continue;
                    }

                    if (!batches.TryGetValue(recipient.Id, out NotificationBatch batch))
                    {
                        batch = new NotificationBatch { Recipient = recipient };
                        batches.Add(recipient.Id, batch);
                    }
                    batch.Results.Add(result);
                    if (!string.IsNullOrEmpty(recipientInfo))
                        batch.RecipientNotes.Add(recipientInfo);
                }
                catch (Exception ex)
                {
                    logError($"Не удалось определить получателя конфигурации [{result.Configuration.SettingName}]: {ex.Message}");
                }
            }
            return batches.Values.ToArray();
        }

        private void SendResultMsg(string moduleName, ExchangeNotificationResult[] results, DBUser recipient, string recipientInfo)
        {
            try
            {
                string configNames = string.Join("; ", results.Select(result => result.Configuration.SettingName));
                if (recipient == null)
                {
                    Module.CurrentLogger.Error($"Не удалось отправить уведомление по конфигурациям [{configNames}]: {recipientInfo}");
                    return;
                }

                int sourceCount = results.Sum(result => result.SourceCount);
                int processedCount = results.Sum(result => result.ProcessedCount);
                bool hasErrors = results.Any(result => result.Failed || result.ProcessedCount < result.SourceCount || result.ProcessedCount == 0);
                string projectNames = string.Join("; ", results.Select(result => result.ProjectName).Distinct());
                string status = hasErrors ? "Отработано с ошибками." : "Отработано без ошибок.";
                string message =
                    $"Модуль: [b]{moduleName}\n[/b]" +
                    $"Анализируемые конфигурации: {configNames}\n" +
                    $"Статус: {status}\n" +
                    $"Метрик производительности: Обработано {processedCount} из {sourceCount} файлов, для проекта: [b]{projectNames}[/b]";
                if (hasErrors)
                    message += $"\nОшибки: См. файл логов у пользователя {SQLiteMainService.CurrentDBUser?.Surname} {SQLiteMainService.CurrentDBUser?.Name}.\n" +
                        $"Путь к логам у пользователя: {Module.CurrentLoggerFullName}";

                if (!string.IsNullOrEmpty(recipientInfo))
                {
                    message += "\nВнимание: " + recipientInfo;
                    Module.CurrentLogger.Warn($"Уведомление по конфигурациям [{configNames}]: {recipientInfo}");
                }

                BitrixMessageSender.SendMsg_ToUser_ByDBUser(recipient, message);
            }
            catch (Exception ex)
            {
                // Ошибка подготовки уведомления не меняет результат экспорта и не мешает другим получателям.
                Module.CurrentLogger.Error($"Не удалось отправить уведомление: {ex.Message}");
            }
        }

        /// <summary>
        /// Недоступного получателя заменяем запускающим пользователем, робота rbim — Тимофеем Куцко.
        /// </summary>
        private static DBUser ResolveNotificationRecipient(int? notificationUserId, out string recipientInfo)
        {
            recipientInfo = string.Empty;
            DBUser recipient = notificationUserId.HasValue
                ? SQLiteMainService.SQLiteUserServiceInst.GetDBUser_ById(notificationUserId.Value)
                : SQLiteMainService.CurrentDBUser;

            if (notificationUserId.HasValue && (recipient == null || recipient.IsFired))
            {
                recipientInfo = $"Получатель уведомления (ID {notificationUserId.Value}) не найден или уволен. " +
                    "Отчёт перенаправлен пользователю, запустившему экспорт. ";
                recipient = SQLiteMainService.CurrentDBUser;
            }

            if (recipient != null && string.Equals(recipient.SystemName, "rbim", StringComparison.OrdinalIgnoreCase))
            {
                recipientInfo += "Получатель — Робот BIM (rbim). Отчёт перенаправлен Тимофею Куцко (tkutsko). ";
                recipient = SQLiteMainService.SQLiteUserServiceInst.GetDBUser_ByUserName("tkutsko");
            }

            if (recipient == null || recipient.IsFired)
            {
                recipientInfo += "Итоговый получатель недоступен; отправка невозможна.";
                return null;
            }

            return recipient;
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
