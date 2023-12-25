using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_Loader.Core.SQLiteData;
using KPLN_Loader.Forms;
using KPLN_Loader.Forms.Common;
using KPLN_Loader.Services;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Loader
{
    public class Application : IExternalApplication
    {
        /// <summary>
        /// Queue для команд на выполнение по событию "OnIdling"
        /// </summary>
        public static Queue<IExecutableCommand> OnIdling_CommandQueue = new Queue<IExecutableCommand>();
        /// <summary>
        /// Основной путь к конфигу БД, которые используются всеми плагинами
        /// </summary>
        public readonly static string SQLMainConfigPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\SQLConfig.json";

        internal delegate void RiseStepProgress(MainStatus mainStatus, string toolTip, System.Windows.Media.Brush brush);
        /// <summary>
        /// Событие, которое посылает сигналы в форму статуса загрузки
        /// </summary>
        internal event RiseStepProgress Progress;

        internal delegate void RiseLoadEvant(LoaderEvantEntity lModule, System.Windows.Media.Brush brush);
        /// <summary>
        /// Событие, которое посылает сигналы в форму статуса загрузки
        /// </summary>
        internal event RiseLoadEvant LoadStatus;

        private readonly static string _diteTime = DateTime.Now.ToString("dd/MM/yyyy_HH/mm/ss");
        private SQLiteService _dbService;
        private EnvironmentService _envService;
        private Logger _logger;
        private const string _ribbonName = "KPLN";
        private string _revitVersion;

        /// <summary>
        /// Кэширование текщего пользователя из БД
        /// </summary>
        public static User CurrentRevitUser { get; private set; }

        /// <summary>
        /// Кэширование текщего отдела
        /// </summary>
        public static SubDepartment CurrentSubDepartment { get; private set; }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentOpened -= new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            LoaderStatusForm loaderStatusForm = new LoaderStatusForm(this);
            loaderStatusForm.Show();

            _revitVersion = application.ControlledApplication.VersionNumber;

            #region Настройка NLog
            LogManager.Setup().LoadConfigurationFromFile(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString() + "\\nlog.config");
            _logger = LogManager.GetLogger("KPLN_Loader");

            string logDirPath = $"c:\\temp\\KPLN_Logs\\{_revitVersion}";
            string logFileName = "KPLN_Loader";
            LogManager.Configuration.Variables["logdir"] = logDirPath;
            LogManager.Configuration.Variables["logfilename"] = logFileName;
            #endregion

            _logger.Info($"Запуск в Revit {_revitVersion}. Версия модуля: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");
            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));
            
            try
            {
                #region Подготовка и проверка окружения
                _envService = new EnvironmentService(_logger, loaderStatusForm, _revitVersion, _diteTime);
                _envService.SQLFilesExistChecker(SQLMainConfigPath);
                _envService.PreparingAndCliningDirectories();

                Progress?.Invoke(MainStatus.Envirnment, "Успешно!", System.Windows.Media.Brushes.Green);
                #endregion

                #region Подготовка/создание пользователя
                string mainDBPath = _envService.DatabasesPaths.FirstOrDefault(d => d.Name.Contains("KPLN_Loader_MainDB")).Path;
                _dbService = new SQLiteService(_logger, mainDBPath);
                CurrentRevitUser = _dbService.Authorization();
                loaderStatusForm.CheckAndSetDebugStatusByUser(CurrentRevitUser);
                CurrentSubDepartment = _dbService.GetSubDepartmentForCurrentUser(CurrentRevitUser);
                loaderStatusForm.UpdateLayout();
                
                // Добавление пользовательской инструкции
                LoaderDescription loaderDescription = _dbService.GetDescriptionForCurrentUser(CurrentRevitUser);
                loaderStatusForm.SetInstruction(loaderDescription);
                loaderStatusForm.LikeStatus += LoaderStatusForm_RiseLikeEvant;

                // Вывод в окно пользователя
                Progress?.Invoke(MainStatus.DbConnection, "Успешно!", System.Windows.Media.Brushes.Green);
                LoadStatus?.Invoke(
                    new LoaderEvantEntity($"Пользователь: [{CurrentRevitUser.Surname} {CurrentRevitUser.Name}], отдел [{CurrentSubDepartment.Code}]"), 
                    System.Windows.Media.Brushes.OrangeRed);
                loaderStatusForm.UpdateLayout();
                #endregion

                // Создание панели Revit
                application.CreateRibbonTab(_ribbonName);

                #region Подготовка, копирование и активация модулей для пользователя
                _logger.Info("Подготовка, копирование библиотек и активация модулей для пользователя:");

                // Коллекция модулей из БД
                IEnumerable<Module> userAllModules = _dbService.GetModulesForCurrentUser(CurrentRevitUser);
                // Подсчет загруженных модулей
                int uploadModules = 0;
                // Флаг на факт загрузки модуля
                bool isModuleLoad = false;
                foreach (Module module in userAllModules)
                {
                    isModuleLoad = false;
                    string moduleVersion = "-";
                    
                    if (module == null)
                    {
                        string msg = $"Модуль/библиотека [{module.Name}] не найден/а!";
                        _logger.Error(msg);
                        LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Red);
                        continue;
                    }
                    
                    try
                    {
                        DirectoryInfo targetDirInfo = _envService.CopyModule(module);
                        String[] spearator = { ".dll" };
                        if (targetDirInfo != null)
                        {
                            // Копирование и активация библиотек (без имплементации IExternalModule)
                            if (module.IsLibraryModule)
                            {
                                foreach (FileInfo file in targetDirInfo.GetFiles())
                                {
                                    if (CheckModuleName(file.Name, spearator))
                                    {
                                        // Каждую dll библиотеки - нужно прогрузить в текущее приложение, чтобы появилась связь в проекте.
                                        // Если этого не сделать - будут проблемы с использованием загружаемых библиотек
                                        System.Reflection.Assembly moduleAssembly = System.Reflection.Assembly.LoadFrom(file.FullName);
                                        FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                                        if (fvi.FileName.Contains("KPLN"))
                                            moduleVersion = fvi.FileVersion;

                                        isModuleLoad = true;
                                    }
                                }

                                // Вывод в окно пользователя и лог
                                string msg = $"Модуль-библиотека [{module.Name}] версии {moduleVersion} успешно скопирован!";
                                _logger.Info(msg);
                                LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Black);
                                uploadModules++;
                            }
                            // Копирование и активация модулей (с имплементацией IExternalModule)
                            else
                            {
                                foreach (FileInfo file in targetDirInfo.GetFiles())
                                {
                                    if (CheckModuleName(file.Name, spearator))
                                    {
                                        System.Reflection.Assembly moduleAssembly = System.Reflection.Assembly.LoadFrom(file.FullName);
                                        Type implemnentationType = moduleAssembly.GetType(file.Name.Split(spearator, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() + ".Module", false);
                                        if (implemnentationType != null)
                                        {
                                            implemnentationType.GetMember("Module");

                                            // Активация
                                            IExternalModule moduleInstance = Activator.CreateInstance(implemnentationType) as IExternalModule;
                                            Result loadingResult = moduleInstance.Execute(application, _ribbonName);
                                            if (loadingResult == Result.Succeeded)
                                            {
                                                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                                                moduleVersion = fvi.FileVersion;

                                                // Вывод в окно пользователя и лог
                                                string msg = $"Модуль [{module.Name}] версии {moduleVersion} успешно активирован!";
                                                _logger.Info(msg);
                                                LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Black);
                                                uploadModules++;
                                                isModuleLoad = true;
                                            }
                                            else
                                            {
                                                // Вывод в окно пользователя и лог
                                                string msg = $"Модуль [{module.Name}] не активирован.";
                                                _logger.Error(msg + $" Проблема: \n{loadingResult}");
                                                LoadStatus?.Invoke(new LoaderEvantEntity(msg + "Подробнее - см. файл логов"), System.Windows.Media.Brushes.Red);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Вывод в окно пользователя и лог
                            string msg = $"Модуль/библиотека {module.Name} - не скопировался по подготовленному пути";
                            _logger.Error(msg);
                            LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Red);
                        }

                        if (!isModuleLoad)
                        {
                            // Вывод в окно пользователя и лог
                            string msg = $"Модуль/библиотека {module.Name} - не активирован. Отсутсвует dll для активации";
                            _logger.Error(msg);
                            LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Red);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Вывод в окно пользователя и лог
                        string msg = $"Локальная ошибка загрузки модуля {module.Name}";
                        _logger.Error(msg + $" \n{ex}");
                        LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Red);
                    }
                }

                // Вывод в окно пользователя
                if (uploadModules == userAllModules.Count())
                    Progress?.Invoke(MainStatus.ModulesActivation, "Успешно!", System.Windows.Media.Brushes.Green);
                else
                    Progress?.Invoke(MainStatus.ModulesActivation, "С замечаниями. Смотри файл логов KPLN_Loader: c:\\temp\\KPLN_Logs\\", System.Windows.Media.Brushes.Orange);
                loaderStatusForm.UpdateLayout();
                #endregion
            }
            catch (Exception ex)
            {
                _logger.Error($"Глобальная ошибка плагина загрузки: \n{ex}");
                _logger.Info($"Инициализация не удалась\n");
                loaderStatusForm.Start_WindowClose();

                return Result.Cancelled;
            }

            _logger.Info($"Успешная инициализация в Revit {_revitVersion}\n");

            #region Подписка на события
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);
            #endregion

            loaderStatusForm.Start_WindowClose();

            return Result.Succeeded;
        }

        /// <summary>
        /// Обработчик события RiseLikeEvant
        /// </summary>
        private void LoaderStatusForm_RiseLikeEvant(int rate, LoaderDescription loaderDescription)
        {
            _dbService.SetLoaderDescriptionUserRank(rate, loaderDescription);
        }

        /// <summary>
        /// Проверка имени dll-модуля на возможность корректной подгрузки
        /// </summary>
        /// <param name="fileName">Имя файла dll-модуля</param>
        /// <param name="spearator">Разделитель имени dll-модуля</param>
        /// <returns></returns>
        private bool CheckModuleName(string fileName, String[] spearator) => fileName.Split(spearator, StringSplitOptions.None).Count() > 1 && !fileName.Contains(".config");

        /// <summary>
        /// Очистка от старых логов
        /// </summary>
        /// <param name="logPath">Путь к логам</param>
        private void ClearingOldLogs(string logPath, string logName)
        {
            if (Directory.Exists(logPath))
            {
                int cnt = 0;
                DirectoryInfo logDI = new DirectoryInfo(logPath);
                foreach (FileInfo log in logDI.EnumerateFiles())
                {
                    if (log.CreationTime.Date < DateTime.Now.AddDays(-5) && log.Name.Contains(logName))
                    {
                        try
                        {
                            log.Delete();
                            cnt++;
                        }
                        // Ошибка будет только если файл занят
                        catch (UnauthorizedAccessException) { }
                        catch (Exception ex)
                        {
                            _logger.Error($"При попытке очистки старых логов произошла ошибка: {ex.Message}");
                        }
                    }
                }
                _logger.Info($"Очистка от старых логов произведена успешно! Удалено {cnt}");
            }
        }

        /// <summary>
        /// Событие возникающее в режимах простоя Revit (когда приложению API становится безопасно обращаться к активному документу)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnIdling(object sender, IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            while (OnIdling_CommandQueue.Count != 0)
            {
                try
                {
                    OnIdling_CommandQueue.Dequeue().Execute(uiapp);
                }
                catch (Exception ex)
                {
                    TaskDialog td = new TaskDialog("ОШИБКА")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = "Произошла серъезная ошибка при выполнении команды, обратись за помощью в BIM-отдел",
                        FooterText = $"См. файлы логов KPLN_Loader: c:\\temp\\KPLN_Logs\\{_revitVersion}"
                    };
                    td.Show();

                    _logger.Error($"Ошибка выполнения в событии OnIdling:\n{ex}");
                }
            }
        }

        /// <summary>
        /// Событие, происходящее при открытии документа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Autodesk.Revit.ApplicationServices.Application app = sender as Autodesk.Revit.ApplicationServices.Application;
            _dbService.SetRevitUserName(app.Username, CurrentRevitUser);
        }
    }
}
