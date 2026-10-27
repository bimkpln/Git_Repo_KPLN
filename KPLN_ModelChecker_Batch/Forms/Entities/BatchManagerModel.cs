using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.ViewModels;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_OpenDocHandler;
using KPLN_Library_OpenDocHandler.Core;
using KPLN_ModelChecker_Batch.Common;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;
using Microsoft.Win32;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using static System.Net.Mime.MediaTypeNames;

namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class BatchManagerModel : INotifyPropertyChanged, IFieldChangedNotifier
    {
        internal const string PluginName = "Пакетная проверка";

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler<FieldChangedEventArgs> FieldChanged;

        /// <summary>
        /// Коллекция проверок для запуска
        /// </summary>
        private readonly CheckEntity[] _checkEntitiesList = new CheckEntity[8]
        {
            new CheckEntity(new CheckLinks()),
            new CheckEntity(new CheckMainLines()),
            new CheckEntity(new CheckFamilies()),
            new CheckEntity(new CheckWorksets()),
            new CheckEntity(new CheckDimensions()),
            new CheckEntity(new CheckMirroredInstances()),
            new CheckEntity(new CheckHoles()),
            new CheckEntity(new CheckFlatsAreaCompare()),
        };

        /// <summary>
        /// Кэширование пути для выбора файлов
        /// </summary>
        private static string _initialDirectoryForOpenFileDialog = @"Y:\";

        private readonly BatchManager _batchManagerWindow;
        private readonly UIApplication _uiapp;
        private string _worksetToCloseNamesStartWith = string.Empty;
        public string _selectedPresets;

        public BatchManagerModel(BatchManager batchManagerWindow, NLog.Logger currentLogger, UIApplication uiapp)
        {
            _uiapp = uiapp;
            _batchManagerWindow = batchManagerWindow;

            CurrentLogger = currentLogger;

            СheckEntities = CollectionViewSource.GetDefaultView(_checkEntitiesList);


            // Добавление в список открытого документа
            Document actDoc = _uiapp.ActiveUIDocument?.Document;
            if (actDoc != null)
            {
                FileEntity fEnt;
                string mPath = SQLiteDocService.GetFileFullName(actDoc);
                if(mPath.IndexOf("rsn", StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    // Инфа по РС от открытой модели собирается налету
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(mPath);
                    RevitServer rs = new RevitServer(modelPath.CentralServerPath, ModuleData.RevitVersion);
                    fEnt = CreateFEnt(mPath, rs);

                }
                else
                    fEnt = CreateFEnt(mPath);
                
                fEnt.IsActiveDoc = true;

                FileEntitiesList.Add(fEnt);
            }


            // Устанавливаю команды
            InfoCommand = new RelayCommand(Info);
            AddFileCommand = new RelayCommand(AddFile);
            AddRSFileCommand = new RelayCommand(AddRSFile);
            DeleteCommand = new RelayCommand(Delete);
            RunCommand = new RelayCommand(Run);
        }

        /// <summary>
        /// Ссылка на коллекцию путей к файлам в конфиге
        /// </summary>
        public ObservableCollection<FileEntity> FileEntitiesList { get; private set; } = new ObservableCollection<FileEntity>();

        /// <summary>
        /// Коллекция проверок для формирования в окне
        /// </summary>
        public ICollectionView СheckEntities { get; set; }

        /// <summary>
        /// Комманда: Получить инфо по элементу
        /// </summary>
        public ICommand InfoCommand { get; }

        /// <summary>
        /// Комманда: Добавить элемент
        /// </summary>
        public ICommand AddFileCommand { get; }

        /// <summary>
        /// Комманда: Добавить элемент с Revit-Server
        /// </summary>
        public ICommand AddRSFileCommand { get; }

        /// <summary>
        /// Комманда: Удалить элемент
        /// </summary>
        public ICommand DeleteCommand { get; }

        /// <summary>
        /// Комманда: Запуск
        /// </summary>
        public ICommand RunCommand { get; }

        /// <summary>
        /// Имя рабочих наборов, которые нужно закрыть (имя начинается с)
        /// </summary>
        public string WorksetToCloseNamesStartWith
        {
            get => _worksetToCloseNamesStartWith;
            set
            {
                _worksetToCloseNamesStartWith = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Фильтрация по открым замечаниям
        /// </summary>
        public string SelectedPresets
        {
            get => _selectedPresets;
            set
            {
                _selectedPresets = value;
                OnPropertyChanged();

                UpdateChecksSelected_ByRBPresets();
            }
        }

        internal static NLog.Logger CurrentLogger { get; set; }

        private static FileEntity CreateFEnt(string mPath, RevitServer revitServer = null)
        {
            string fileSize;
            if (revitServer != null)
            {
                string rsHost = revitServer.Host;
                string rsRelativePath = mPath.Substring($"RSN:\\\\{rsHost}".Length).TrimStart('\\');
                // mPath = "RSN:\\host\Folder\Sub\Model.rvt"  ->  нужно "Folder\Sub\Model.rvt"
                fileSize = (revitServer.GetModelInfo(rsRelativePath).ModelSize / 1024 / 1024).ToString();
            }
            else
                fileSize = (new FileInfo(mPath).Length / 1024 / 1024).ToString();

            string fileName;
            if(mPath.Contains('\\'))
                fileName = mPath.Split('\\').LastOrDefault();
            else
                fileName = mPath.Split('/').LastOrDefault();

            return new FileEntity(fileName, mPath, fileSize);
        }

        /// <summary>
        /// Обновить список выбранных проверок согласно быстрому выбору
        /// </summary>
        private void UpdateChecksSelected_ByRBPresets()
        {
            if (SelectedPresets == "Preset_All")
                СheckEntities = CollectionViewSource.GetDefaultView(
                    _checkEntitiesList.Select(e => e.IsChecked = true));
            else if (SelectedPresets == "Preset_AR")
                СheckEntities = CollectionViewSource.GetDefaultView(
                    _checkEntitiesList.Select(e => e.IsChecked = true));
            else
            {
                // Проверки площадей АР
                string checkFlatsAreaCompareName = new CheckFlatsAreaCompare().PluginName;
                // Проверки отверстий АР
                string checkHolesName = new CheckHoles().PluginName;
                // Проверки зеркальных
                string checkMirroredInstancesName = new CheckMirroredInstances().PluginName;


                List<string> removeCheckNames = new List<string>();
                if (SelectedPresets == "Preset_OV")
                {
                    removeCheckNames.Add(checkFlatsAreaCompareName);
                    removeCheckNames.Add(checkHolesName);
                }
                else if (SelectedPresets == "Preset_KR"
                    || SelectedPresets == "Preset_VK"
                    || SelectedPresets == "Preset_SS"
                    || SelectedPresets == "Preset_EOM")
                {
                    removeCheckNames.Add(checkMirroredInstancesName);

                    removeCheckNames.Add(checkFlatsAreaCompareName);
                    removeCheckNames.Add(checkHolesName);
                }

                // Чистим коллекцию в зависимости от исключений
                List<CheckEntity> filteredEnt = new List<CheckEntity>();
                foreach (CheckEntity ent in _checkEntitiesList)
                {
                    if (removeCheckNames.Contains(ent.Name))
                    {
                        ent.IsChecked = false;
                        filteredEnt.Add(ent);
                    }
                    else
                    {
                        ent.IsChecked = true;
                        filteredEnt.Add(ent);
                    }
                }
            }
        }

        private void Info(object parameter)
        {
            // Получаем выбранные элементы
            if (parameter is System.Collections.IList list)
            {
                if (list.Count != 1)
                    MessageBox.Show(_batchManagerWindow, "Информацию можно смотерть только по отдельным файлам", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                {
                    var itemsToInfo = list.Cast<FileEntity>().ToList();
                    foreach (var item in itemsToInfo)
                    {
                        UserStringInfo userStringInput = new UserStringInfo(true, item.Name, item.Path);
                        userStringInput.ShowDialog();
                    }
                }
            }
        }

        private void Delete(object parameter)
        {
            // Получаем выбранные элементы
            if (parameter is System.Collections.IList list)
            {
                var td = MessageBox.Show(_batchManagerWindow, $"Сейчас будет удалены элементы в количестве \"{list.Count}\" шт.", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (td == MessageBoxResult.Yes)
                {
                    var itemsToDelete = list.Cast<FileEntity>().ToList();
                    // Удаляем каждый выбранный элемент из коллекции
                    foreach (var item in itemsToDelete)
                    {
                        FileEntitiesList.Remove(item);
                    }
                }
            }
        }

        private void AddFile(object parameter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Выберите Revit-файлы",
                Multiselect = true,
                Filter = "Revit Files (*.rvt)|*.rvt",
                InitialDirectory = _initialDirectoryForOpenFileDialog,
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _initialDirectoryForOpenFileDialog = openFileDialog.FileName;
                foreach (string filePath in openFileDialog.FileNames)
                {
                    AddToFileEntitiesWithCheck(CreateFEnt(filePath));
                }
            }
        }

        private void AddRSFile(object parameter)
        {
            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_batchManagerWindow, ModuleData.RevitVersion);
            if (rsFilesPickForm == null)
                return;

            if ((bool)rsFilesPickForm.ShowDialog())
            {
                foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
                {
                    string filePath = $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}";
                    AddToFileEntitiesWithCheck(CreateFEnt(filePath, SelectFilesFromRevitServer.CurrentRevitServer));
                }
            }
        }

        /// <summary>
        /// Добавить эл-т в коллекцию FileEntitiesList с предпроверкой
        /// </summary>
        private void AddToFileEntitiesWithCheck(FileEntity entity)
        {
            if (!FileEntitiesList.Any(listEnt => listEnt.Path.Equals(entity.Path) || listEnt.Name.Equals(entity.Name)))
                FileEntitiesList.Add(entity);
            else
                MessageBox.Show(_batchManagerWindow, "Файл по указнному пути или с тем же именем уже есть в списке. В добавлении отказано", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void Run(object parameter)
        {
            string excelPath = ExportToExcel.SetPath();
            if (string.IsNullOrEmpty(excelPath)) return;

            bool hasErrors = false;
            FileEntity[] files = FileEntitiesList.ToArray();
            CheckEntity[] checks = _checkEntitiesList.Where(check => check.IsChecked).ToArray();
            if (files.Length == 0 || checks.Length == 0)
            {
                MessageBox.Show("Выберите хотя бы одну модель и одну проверку", PluginName);
                return;
            }

            List<string> reportPaths = new List<string>();
            int completedModels = 0;
            bool cancelled = false;
            ProgressInfoViewModel progress = new ProgressInfoViewModel
            {
                ProcessTitle = "Модели",
                MaxProgress = files.Length,
                HasSecondaryProgress = true,
                SecondaryProcessTitle = "Проверки текущей модели",
                SecondaryMaxProgress = checks.Length,
                WaitForCancellationCompletion = true,
            };
            ProgressWindow progressWindow = ProgressWindow.ShowInBackground(
                progress, "KPLN: Пакетная проверка", false, _uiapp.MainWindowHandle);

            try
            {

                CurrentLogger.Info($"Старт плагина \"{PluginName}\".");
                using (UIContrAppSubscriber subscriber = new UIContrAppSubscriber(ModuleData.RevitUIControlledApp, CurrentLogger, this))
                {
                    foreach (FileEntity file in files)
                    {
                        if (progress.IsCancellationRequested)
                        {
                            cancelled = true;
                            break;
                        }

                        progress.CurrentEventName = file.Name;
                        progress.SecondaryCurrentProgress = 0;
                        progress.SecondaryEventName = "Ожидание открытия модели";
                        Document doc = null;
                        ExcelDataEntity excelEnt = new ExcelDataEntity()
                        {
                            CheckRunData = DateTime.Now.ToString("G"),
                            FileName = file.Name,
                            FileSize = file.FileSize,
                            ChecksData = new List<CheckData>(),
                        };

                        // Ошибка одной модели не прерывает обработку следующих.
                        try
                        {
                            progress.Report(completedModels, files.Length, "Открытие модели");
                            CurrentLogger.Info($"Начинаю проверку файла: \"{file.Name}\"");
                            ModelPath docModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(file.Path);
                            // Открываю документ, если нужно
                            if (!file.IsActiveDoc)
                            {
                                #region Подготовка OpenOptions
                                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(docModelPath);
                                IList<WorksetId> worksetIds = new List<WorksetId>();

                                int openWS = 0;
                                int allWS = 0;
                                foreach (WorksetPreview worksetPrev in worksets)
                                {
                                    allWS++;
                                    if (WorksetToCloseNamesStartWith.Count() == 0)
                                    {
                                        openWS++;
                                        worksetIds.Add(worksetPrev.Id);
                                    }
                                    else if (!WSName_IsMatchByRules(worksetPrev.Name, WorksetToCloseNamesStartWith))
                                    {
                                        openWS++;
                                        worksetIds.Add(worksetPrev.Id);
                                    }
                                }
                                excelEnt.CountedWorksets = allWS;
                                excelEnt.OpenedWorksets = openWS;

                                OpenOptions openOptions = new OpenOptions()
                                {
                                    DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                                };

                                WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                                openConfig.Open(worksetIds);
                                openOptions.SetOpenWorksetsConfiguration(openConfig);
                                #endregion

                                // Добавил задержку, т.к. бывает файл не хочет открыться, и ошибка "was thrown by Revit or by one of its external applications"
                                for (int delay = 0; delay < 20; delay++)
                                {
                                    progress.ThrowIfCancellationRequested();
                                    Thread.Sleep(100);
                                }
                                progress.ThrowIfCancellationRequested();

                                doc = _uiapp.Application.OpenDocumentFile(docModelPath, openOptions);
                            }
                            else
                                doc = _uiapp.ActiveUIDocument.Document;


                            progress.ThrowIfCancellationRequested();

                            // Обрабатываю документ
                            var linkTypes = new FilteredElementCollector(doc)
                                .OfClass(typeof(RevitLinkType))
                                .Cast<RevitLinkType>()
                                .ToArray();

                            int modelLinks = linkTypes.Length;
                            int loadedLinks = linkTypes.Where(rlt => rlt.GetLinkedFileStatus() == LinkedFileStatus.Loaded).Count();
                            excelEnt.CountedLinks = modelLinks;
                            excelEnt.OpenedLinks = loadedLinks;


                            // Изолируем подготовку и выполнение каждой выбранной проверки.
                            int completedChecks = 0;
                            foreach (CheckEntity check in checks)
                            {
                                progress.ReportSecondary(completedChecks, checks.Length, check.Name);
                                progress.MaintStatus = "Выполняется проверка";
                                CheckData result = new CheckData()
                                {
                                    PluginName = check.Name,
                                    CheckRunResult = CheckResultStatus.Failed,
                                };

                                try
                                {
                                    CurrentLogger.Info($"Начинаю проверку: \"{check.Name}\"");
                                    // Новая проверка не наследует частичные данные предыдущей модели.
                                    AbstrCheck currentCheck = (AbstrCheck)Activator.CreateInstance(check.CurrentAbstrCheck.GetType());
                                    currentCheck.Set_UIAppData(_uiapp, doc);
                                    progress.ThrowIfCancellationRequested();
                                    Element[] elements = currentCheck.GetElemsToCheck();
                                    progress.ThrowIfCancellationRequested();
                                    result.CheckRunResult = currentCheck.ExecuteCheck(elements, false);
                                    result.PluginCheckerEntitiesColl = currentCheck.CheckerEntitiesColl;

                                    if (result.CheckRunResult == CheckResultStatus.Failed)
                                    {
                                        hasErrors = true;
                                        result.ErrorMessage = "Проверка завершилась с ошибкой. Подробности см. в окне ошибок проверки.";
                                        ReportBatchError(file.Name, check.Name, result.ErrorMessage);
                                    }
                                }
                                catch (OperationCanceledException) when (progress.IsCancellationRequested)
                                {
                                    throw;
                                }
                                catch (Exception ex)
                                {
                                    hasErrors = true;
                                    result.CheckRunResult = CheckResultStatus.Failed;
                                    result.ErrorMessage = ex.Message;
                                    ReportBatchError(file.Name, check.Name, ex.ToString());
                                }

                                excelEnt.ChecksData.Add(result);
                                progress.SecondaryCurrentProgress = ++completedChecks;
                                progress.ThrowIfCancellationRequested();
                            }
                        }
                        catch (OperationCanceledException) when (progress.IsCancellationRequested)
                        {
                            cancelled = true;
                        }
                        catch (Exception ex)
                        {
                            hasErrors = true;
                            excelEnt.ChecksData.Add(new CheckData()
                            {
                                PluginName = "Открытие / подготовка модели",
                                CheckRunResult = CheckResultStatus.Failed,
                                ErrorMessage = ex.Message,
                            });
                            ReportBatchError(file.Name, "Открытие / подготовка модели", ex.ToString());
                        }
                        finally
                        {
                            // Отмена не превращает незапущенные проверки в успешные.
                            if (progress.IsCancellationRequested)
                            {
                                cancelled = true;
                                foreach (CheckEntity pending in checks.Where(check => !excelEnt.ChecksData.Any(data => data.PluginName == check.Name)))
                                    excelEnt.ChecksData.Add(new CheckData
                                    {
                                        PluginName = pending.Name,
                                        CheckRunResult = CheckResultStatus.Cancelled,
                                        ErrorMessage = "Не запущена: пакетная проверка отменена пользователем",
                                    });
                            }
                            else
                                progress.MaintStatus = "Сохранение отчёта и закрытие модели";

                            // Сохраняем результат до закрытия документа, пока доступны его элементы.
                            try
                            {
                                string reportPath = ExportToExcel.Run(excelPath, new[] { excelEnt });
                                reportPaths.Add(reportPath);
                                CurrentLogger.Info($"Сохранён отчёт: {reportPath}");
                            }
                            catch (Exception ex)
                            {
                                hasErrors = true;
                                ReportBatchError(file.Name, "Сохранение Excel", ex.ToString());
                            }
                            finally
                            {
                                if (!file.IsActiveDoc && doc != null)
                                {
                                    try
                                    {
                                        doc.Close(false);
                                    }
                                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                    {
                                        // Документ может быть связью открытой модели.
                                    }
                                    catch (Exception ex)
                                    {
                                        hasErrors = true;
                                        ReportBatchError(file.Name, "Закрытие модели", ex.ToString());
                                    }
                                }
                            }
                        }
                        if (cancelled)
                        {
                            if (excelEnt.ChecksData.Count == checks.Length
                                && excelEnt.ChecksData.All(data => data.CheckRunResult != CheckResultStatus.Cancelled))
                                progress.CurrentProgress = ++completedModels;
                            break;
                        }

                        progress.CurrentProgress = ++completedModels;
                    }
                }
            }
            finally
            {
                cancelled |= progress.IsCancellationRequested;
                progress.IsCancellationEnabled = false;
                progressWindow.CloseSafely();
            }

            if (cancelled)
            {
                string message = $"Пакетная проверка отменена. Обработано моделей: {completedModels}/{files.Length}. " +
                    "Следующие модели не запускались. Сохранённые отчёты перечислены ниже.\n" +
                    (reportPaths.Count == 0 ? "Отчёты не сохранены" : string.Join("\n", reportPaths));
                CurrentLogger.Info(message);
                HtmlOutput.Print(message, MessageType.Warning);
                return;
            }

            string fileNames = string.Join("; ", files.Select(file => file.Name));
            string checkNames = string.Join("; ", checks.Select(check => check.Name));
            string reports = reportPaths.Count == 0 ? "Отчёты не сохранены" : string.Join("\n", reportPaths);
            if (hasErrors)
                SendResultErrorMsg(fileNames, checkNames, reports);
            else
                SendResultMsg(fileNames, checkNames, reports);
        }

        private static void ReportBatchError(string fileName, string stage, string details)
        {
            string message = $"Модель: {fileName}. Этап: {stage}. Ошибка: {details}";
            // Сбой вывода сообщения тоже не должен прерывать пакет.
            try
            {
                CurrentLogger.Error(message);
            }
            catch (Exception)
            {
                System.Diagnostics.Trace.WriteLine(message);
            }

            try
            {
                HtmlOutput.Print(message, MessageType.Error);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex);
            }
        }

        /// <summary>
        /// Поиск части имени в РН. РФ файл - ВСЕГДА ОТКРЫВАЕМ
        /// </summary>
        /// <returns></returns>
        private static bool WSName_IsMatchByRules(string rnName, string ruleString)
        {
            // Нет исключений. Анализ прекращён
            if (string.IsNullOrWhiteSpace(rnName) || string.IsNullOrWhiteSpace(ruleString))
                return false;

            // РН с РФ всегда должен открыться
            if (rnName.IndexOf("рф", StringComparison.OrdinalIgnoreCase) >= 0
                || rnName.IndexOf("разб", StringComparison.OrdinalIgnoreCase) >= 0)
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
        /// Отправка результата пользователю в месенджер
        /// </summary>
        private static void SendResultMsg(string fileNames, string checkNames, string excelFilePath)
        {
            BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                SQLiteMainService.CurrentDBUser,
                $"Модуль: [b]{PluginName}\n[/b]" +
                $"Анализируемые файлы: {fileNames}\n" +
                $"Файлы с результатами проверок: {excelFilePath}\n" +
                $"Запускаемые проверки: {checkNames}\n" +
                $"Статус: Отработано. Глобальных ошибок нет. Возможны локальные, внимательно изучи полученный файл с результатом проверок.\n");
        }

        /// <summary>
        /// Отправка результата об ошибке пользователю в месенджер
        /// </summary>
        private static void SendResultErrorMsg(string fileNames, string checkNames, string excelFilePath)
        {
            BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                SQLiteMainService.CurrentDBUser,
                $"Модуль: [b]{PluginName}\n[/b]" +
                $"Анализируемые файлы: {fileNames}\n" +
                $"Файлы с результатами проверок: {excelFilePath}\n" +
                $"Запускаемые проверки: {checkNames}\n" +
                $"Статус: Отработано с глобальными ошибками.\n" +
                $"Ошибки: См. файл логов у пользователя {SQLiteMainService.CurrentDBUser.Surname} {SQLiteMainService.CurrentDBUser.Name}.\n" +
                $"Путь к логам у пользователя: C:\\KPLN_Temp\\KPLN_Logs\\{ModuleData.RevitVersion}");
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void OnFieldChanged(FieldChangedEventArgs e) =>
            FieldChanged?.Invoke(this, e);
    }
}
