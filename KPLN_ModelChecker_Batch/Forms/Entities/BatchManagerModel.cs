using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_OpenDocHandler;
using KPLN_Library_OpenDocHandler.Core;
using KPLN_Library_PluginActivityWorker;
using KPLN_Library_SQLiteWorker;
using KPLN_ModelChecker_Batch.Common;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Commands;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Input;

namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class BatchManagerModel : INotifyPropertyChanged, IFieldChangedNotifier
    {
        internal const string PluginName = "Пакетная\nпроверка";

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler<FieldChangedEventArgs> FieldChanged;

        /// <summary>
        /// Кэширование пути для выбора файлов
        /// </summary>
        private static string _initialDirectoryForOpenFileDialog = @"Y:\";

        private readonly UIApplication _uiapp;
        private string _worksetToCloseNamesStartWith = string.Empty;

        public BatchManagerModel(NLog.Logger currentLogger, UIApplication uiapp)
        {
            CurrentLogger = currentLogger;
            _uiapp = uiapp;

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

        public ObservableCollection<CheckEntity> CheckEntitiesList { get; private set; } = new ObservableCollection<CheckEntity>()
        {
            new CheckEntity(new CheckLinks()),
            new CheckEntity(new CheckMainLines()),
            new CheckEntity(new CheckWorksets()),
        };

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

        internal static NLog.Logger CurrentLogger { get; set; }

        private void Info(object parameter)
        {
            // Получаем выбранные элементы
            if (parameter is System.Collections.IList list)
            {
                if (list.Count != 1)
                {
                    UserDialog cd = new UserDialog("Предупреждение", $"Информацию можно смотерть только по отдельным файлам");
                    cd.ShowDialog();
                }
                else
                {
                    var itemsToDelete = list.Cast<FileEntity>().ToList();
                    foreach (var item in itemsToDelete)
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
                UserDialog cd = new UserDialog("Предупреждение", $"Сейчас будет удалены элементы в количестве {list.Count} шт.");
                if ((bool)cd.ShowDialog())
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
                    string fileName = Path.GetFileName(filePath);
                    AddToFileEntitiesWithCheck(new FileEntity(fileName, filePath));
                }
            }
        }

        private void AddRSFile(object parameter)
        {
            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(ModuleData.RevitVersion);
            if (rsFilesPickForm == null)
                return;

            if ((bool)rsFilesPickForm.ShowDialog())
            {
                foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
                {
                    AddToFileEntitiesWithCheck(new FileEntity(formEntity.Name, $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}"));
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
            {
                UserDialog cd = new UserDialog("Предупреждение", $"Файл по указнному пути или с тем же именем уже есть в списке. В добавлении отказано");
                cd.ShowDialog();
            }
        }

        private void Run(object parameter)
        {
            bool isFileError = false;
            bool isCheckError = false;
            string fileNames = string.Join("; ", FileEntitiesList.Select(ent => ent.Name));
            string checkNames = string.Join("; ", CheckEntitiesList.Select(ent => ent.Name));
            List<ExcelDataEntity> excelDataEntities = new List<ExcelDataEntity>();

            string excelPath = ExportToExcel.SetPath();
            if (string.IsNullOrEmpty(excelPath)) return;

            CurrentLogger.Info($"Старт палгина \"{PluginName}.\"");
            using (UIContrAppSubscriber subscriber = new UIContrAppSubscriber(ModuleData.RevitUIControlledApp, CurrentLogger, this))
            {
                foreach (FileEntity file in FileEntitiesList)
                {
                    CurrentLogger.Info($"Начинаю проверку файла: \"{file.Name}\"");

                    ExcelDataEntity excelEnt = new ExcelDataEntity()
                    {
                        CheckRunData = DateTime.Now.ToString("G"),
                        FileName = file.Name,
                        FileSize = (new FileInfo(file.Path).Length / 1024 / 1024).ToString(),
                        ChecksData = new List<CheckData>(),
                    };

                    #region Открытие документа
                    ModelPath docModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(file.Path);

                    #region Подготовка OpenOptions
                    IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(docModelPath);
                    IList<WorksetId> worksetIds = new List<WorksetId>();

                    string[] wsExceptions = WorksetToCloseNamesStartWith.Split('~');
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
                        else if (!wsExceptions.Any(name => worksetPrev.Name.StartsWith(name)))
                        {
                            openWS++;
                            worksetIds.Add(worksetPrev.Id);
                        }
                    }
                    excelEnt.OpenedWorksets = $"{openWS}/{allWS}";

                    OpenOptions openOptions = new OpenOptions()
                    {
                        DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                    };

                    WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                    openConfig.Open(worksetIds);
                    openOptions.SetOpenWorksetsConfiguration(openConfig);
                    #endregion

                    Document doc = null;
                    // Открываем документ по указанному пути
                    try
                    {
                        // Добавил задержку, т.к. бывает файл не хочет открыться, и ошибка "was thrown by Revit or by one of its external applications"
                        Thread.Sleep(2000);

                        doc = _uiapp.Application.OpenDocumentFile(docModelPath, openOptions);

                        var linkTypes = new FilteredElementCollector(doc)
                            .OfClass(typeof(RevitLinkType))
                            .Cast<RevitLinkType>()
                            .ToArray();

                        int modelLinks = linkTypes.Length;
                        int loadedLinks = linkTypes.Where(rlt => rlt.GetLinkedFileStatus() == LinkedFileStatus.Loaded).Count();
                        excelEnt.OpenedLinks = $"{loadedLinks}/{modelLinks}";

                    }
                    catch (Autodesk.Revit.Exceptions.FileNotFoundException)
                    {
                        string modelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
                        string msg = $"Путь к файлу {modelPath} - не существует. Внимательно проверь путь и наличие модели по указанному пути";
                        CurrentLogger.Error(msg);
                        doc.Close(false);

                        continue;
                    }
                    catch (Exception ex)
                    {
                        CurrentLogger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath)}). Нужно вмешаться человеку, " +
                            $"ошибка при открытии: \"{ex.Message}\"");
                        isFileError = true;
                        doc.Close(false);

                        continue;
                    }
                    #endregion

                    #region Запуск проверки
                    string checkName = string.Empty;
                    try
                    {
                        foreach (CheckEntity check in CheckEntitiesList)
                        {
                            if (!check.IsChecked) continue;

                            checkName = check.Name;
                            CurrentLogger.Info($"Начинаю проверку с именем: \"{checkName}\"");

                            CheckerEntity[] docErrors = check.RunCommand(doc);

                            if (docErrors != null && docErrors.Length > 0)
                            {
                                excelEnt.ChecksData.Add(new CheckData()
                                {
                                    PluginName = checkName,
                                    CheckEntities = docErrors
                                });
                            }

                            CurrentLogger.Info($"Завершена проверка с именем: \"{checkName}\"");
                        }
                    }
                    catch (Exception ex)
                    {
                        isCheckError = true;
                        CurrentLogger.Error($"Работа плагина \"{PluginName}\" с проверкой прервана\"{checkName}\". Ошибка: \"{ex.Message}\"\n");
                    }


                    doc.Close(false);
                    #endregion

                    excelDataEntities.Add(excelEnt);

                    CurrentLogger.Info($"Завершена проверка файла: \"{file.Name}\"");
                }

                CurrentLogger.Info($"Плагин \"{PluginName}\" завершил работу.");
            }


            ExportToExcel.Run(excelPath, excelDataEntities.ToArray());


            if (isCheckError || isFileError)
                SendResultErrorMsg(fileNames, checkNames);
            else
                SendResultMsg(fileNames, checkNames);

            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void OnFieldChanged(FieldChangedEventArgs e) =>
            FieldChanged?.Invoke(this, e);

        /// <summary>
        /// Отправка результата пользователю в месенджер
        /// </summary>
        private void SendResultMsg(string fileNames, string checkNames)
        {
            BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                DBMainService.CurrentDBUser,
                $"Модуль: [b]{PluginName}\n[/b]" +
                $"Анализируемые файлы: {fileNames}\n" +
                $"Запускаемые проверки: {checkNames}\n" +
                $"Статус: Отработано без ошибок.\n");
        }

        /// <summary>
        /// Отправка результата об ошибке пользователю в месенджер
        /// </summary>
        private void SendResultErrorMsg(string fileNames, string checkNames)
        {
            BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                DBMainService.CurrentDBUser,
                $"Модуль: [b]{PluginName}\n[/b]" +
                $"Анализируемые файлы: {fileNames}\n" +
                $"Запускаемые проверки: {checkNames}\n" +
                $"Статус: Отработано с ошибками.\n" +
                $"Ошибки: См. файл логов у пользователя {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name}.\n" +
                $"Путь к логам у пользователя: C:\\KPLN_Temp\\KPLN_Logs\\{ModuleData.RevitVersion}");
        }
    }
}
