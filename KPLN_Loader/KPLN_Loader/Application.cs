using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms;
using KPLN_Loader.Forms.Common;
using KPLN_Loader.Services;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Loader
{
    public class Application : IExternalApplication
    {
        /// <summary>
        /// Кэширую версию используемого Revit
        /// </summary>
        public static string RevitVersion { get; private set; }

        /// <summary>
        /// Путь к папке с файлами для кэширования (временные конфиги, логи сессий и т.п.)
        /// </summary>
        public static string MainCashFolder { get; private set; }

        /// <summary>
        /// Кэширование текщего пользователя из БД
        /// </summary>
        public static User CurrentRevitUser { get; private set; }

        /// <summary>
        /// Кэширование текщего отдела
        /// </summary>
        public static SubDepartment CurrentSubDepartment { get; private set; }

        /// <summary>
        /// Queue для команд на выполнение по событию "OnIdling"
        /// </summary>
        public static Queue<IExecutableCommand> OnIdling_CommandQueue = new Queue<IExecutableCommand>();

        /// <summary>
        /// ВНУТРИ KPLN: Основной путь к основным конфигам для работы всей экосистемы
        /// </summary>
        public readonly static string MainConfigPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_Config.json";

        /// <summary>
        /// ДЛЯ СУБЧИКА: ID листа гугл таблицы, которая выступает в роли БД KPLN_Loader
        /// </summary>
        public readonly static string GTab_KPLNLoader_SheetId = "1sFx8Vd_n9RI9rNFUjtiJcfGK1rJo8v4Bb993dnrub-I";

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
        private static string _pahtToLoaderDll;
        private const string _ribbonName = "KPLN";
        private readonly List<IExternalModule> _moduleInstances = new List<IExternalModule>();
        private SQLiteService _dbService;
        private GTabService _gtabService;
        private EnvironmentService _envService;
        private Logger _logger;

        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= new EventHandler<IdlingEventArgs>(OnIdling);
            application.ControlledApplication.DocumentOpened -= new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);

            foreach(IExternalModule module in _moduleInstances)
            {
                module.Close();
            }
            
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Создание панели Revit
            application.CreateRibbonTab(_ribbonName);
            
            RevitVersion = application.ControlledApplication.VersionNumber;
            _pahtToLoaderDll = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            #region Настройка NLog
            // Данный логгер должен содержать все настройки для каждого отдельного плагина. Это связано с инициализацией dll самим ревитом.
            LogManager.Setup().LoadConfigurationFromFile(_pahtToLoaderDll + "\\nlog.config");
            _logger = LogManager.GetLogger("KPLN_Loader");

            string windrive = Path.GetPathRoot(Environment.SystemDirectory);
            MainCashFolder = $"{windrive}KPLN_Temp";

            string logDirPath = $"{MainCashFolder}\\KPLN_Logs\\{RevitVersion}";
            string logFileName = "KPLN_Loader";
            LogManager.Configuration.Variables["loader_logdir"] = logDirPath;
            LogManager.Configuration.Variables["loader_logfilename"] = logFileName;
            #endregion
            
            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));

            _logger.Info($"Запуск в Revit {RevitVersion}. Версия модуля: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");

            // Определяю сеть и определяю сценарии загрузки
            string domainName = IPGlobalProperties.GetIPGlobalProperties().DomainName;
            // Для KPLN (с копированием модулей и стартовым окном)
            if (domainName.Equals("stinproject.local"))
            {
                LoaderStatusForm loaderStatusForm = null;
                ManualResetEvent formReady = new ManualResetEvent(false);
                Thread uiThread = new Thread(() =>
                {
                    loaderStatusForm = new LoaderStatusForm(this);
                    loaderStatusForm.Show();
                    formReady.Set();
                    System.Windows.Threading.Dispatcher.Run();
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.IsBackground = true;
                uiThread.Start();
                formReady.WaitOne();

                try
                {
                    #region Подготовка и проверка окружения
                    _envService = new EnvironmentService(_logger, loaderStatusForm, RevitVersion, _diteTime)
                        .ConfigFileChecker()
                        .PreparingAndCliningDirectories();

                    Progress?.Invoke(MainStatus.Envirnment, "Успешно!", System.Windows.Media.Brushes.Green);
                    #endregion

                    #region Подготовка/создание пользователя
                    
                    string mainDBPath = _envService.DatabaseConfigs.FirstOrDefault(d => d.Name == EnvironmentService.DatabaseConfigs_LoaderMainDB).Path;
                    _dbService = new SQLiteService(_logger, mainDBPath);
                    CurrentRevitUser = _dbService.Authorization(_envService);
                    if (CurrentRevitUser == null)
                    {
                        Progress?.Invoke(MainStatus.DbConnection, "Критическая ошибка пользователя! Подробнее - см. файл логов", System.Windows.Media.Brushes.Red);
                        return Result.Cancelled;
                    }

                    loaderStatusForm.SetDebugStatus(CurrentRevitUser.IsDebugMode);
                    bool isUserDataUpdated = _dbService.SetUserLastConnectionDate(CurrentRevitUser);
                    CurrentSubDepartment = _dbService.GetSubDepartmentForCurrentUser(CurrentRevitUser);
                    loaderStatusForm.Dispatcher.Invoke(() => loaderStatusForm.UpdateLayout());

                    // Добавление пользовательской инструкции
                    LoaderDescription loaderDescription = _dbService.GetDescriptionForCurrentUser(CurrentRevitUser);
                    loaderStatusForm.SetInstruction(loaderDescription);
                    loaderStatusForm.LikeStatus += LoaderStatusForm_RiseLikeEvant;

                    if (isUserDataUpdated && CurrentSubDepartment != null)
                        Progress?.Invoke(MainStatus.DbConnection, "Успешно!", System.Windows.Media.Brushes.Green);
                    else
                        Progress?.Invoke(MainStatus.DbConnection, "Замечания при подключения к БД! Подробнее - см. файл логов", System.Windows.Media.Brushes.Orange);
                    
                    // Вывод в окно пользователя
                    LoadStatus?.Invoke(
                        new LoaderEvantEntity($"Пользователь: [{CurrentRevitUser.Surname} {CurrentRevitUser.Name}], отдел [{CurrentSubDepartment.Code}]"),
                        System.Windows.Media.Brushes.OrangeRed);
                    loaderStatusForm.Dispatcher.Invoke(() => loaderStatusForm.UpdateLayout());
                    #endregion

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
                            // Копирование модуля
                            DirectoryInfo targetDirInfo = _envService.CopyModule(module, CurrentRevitUser.IsDebugMode);

                            String[] spearator = { ".dll" };
                            if (targetDirInfo != null)
                            {
                                // Активация библиотеки (без имплементации IExternalModule)
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
                                            if (moduleAssembly.FullName.Contains("KPLN"))
                                                moduleVersion = fvi.FileVersion;

                                            isModuleLoad = true;
                                        }
                                    }

                                    if (isModuleLoad)
                                    {
                                        // Вывод в окно пользователя и лог
                                        string msg = $"Модуль-библиотека [{module.Name}] версии {moduleVersion} успешно скопирован!";
                                        _logger.Info(msg);
                                        LoadStatus?.Invoke(new LoaderEvantEntity(msg), System.Windows.Media.Brushes.Black);
                                        uploadModules++;
                                    }
                                }
                                // Ативация модуля (с имплементацией IExternalModule)
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
                                                    _moduleInstances.Add(moduleInstance);
                                                    
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
                        Progress?.Invoke(MainStatus.ModulesActivation, $"С замечаниями. Смотри файл логов KPLN_Loader: {MainCashFolder}\\KPLN_Logs\\", System.Windows.Media.Brushes.Orange);
                    loaderStatusForm.Dispatcher.Invoke(() => loaderStatusForm.UpdateLayout());
                    #endregion
                }
                catch (Exception ex)
                {
                    _logger.Error($"Глобальная ошибка плагина загрузки: \n{ex}");
                    _logger.Info($"Инициализация не удалась\n");
                    loaderStatusForm.Dispatcher.Invoke(() => loaderStatusForm.Close());

                    Exception currentEx = ex.InnerException ?? ex;
                    MessageBox.Show(
                        $"Инициализация не удалась. Отправь в BIM-отдел KPLN: {currentEx.Message}",
                        "Ошибка KPLN",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);

                    return Result.Cancelled;
                }

                loaderStatusForm.Start_WindowClose();
            }
            // Для всех остальных
            else
            {
                #region Подготовка/создание пользователя
                _gtabService = new GTabService(_logger, GTab_KPLNLoader_SheetId, RevitVersion);
                CurrentRevitUser = _gtabService.Authorization();
                if (CurrentRevitUser == null)
                    return Result.Cancelled;

                CurrentSubDepartment = _gtabService.GetSubDepartmentForCurrentUser(CurrentRevitUser);
                #endregion

                // Активация модулей для пользователя
                _logger.Info("Активация модулей для пользователя:");
                IEnumerable<Module> userAllModules = ExtraNetUserModules(application);
            }

            _logger.Info($"Успешная инициализация в Revit {RevitVersion}\n");

            #region Подписка на события
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);
            #endregion


            return Result.Succeeded;
        }

        /// <summary>
        /// Обработчик события RiseLikeEvant
        /// </summary>
        private void LoaderStatusForm_RiseLikeEvant(int rate, LoaderDescription loaderDescription) => _dbService.SetLoaderDescriptionUserRank(rate, loaderDescription);

        /// <summary>
        /// Проверка имени dll-модуля на возможность корректной подгрузки
        /// </summary>
        /// <param name="fileName">Имя файла dll-модуля</param>
        /// <param name="spearator">Разделитель имени dll-модуля</param>
        /// <returns></returns>
        private bool CheckModuleName(string fileName, String[] spearator) => fileName.Split(spearator, StringSplitOptions.None).Count() > 1 && !fileName.Contains(".config");

        /// <summary>
        /// Получить коллекцию модулей для пользователей типа ExtraNet (субчики)
        /// </summary>
        /// <returns></returns>
        internal IEnumerable<Module> ExtraNetUserModules(UIControlledApplication application)
        {
            List<Module> modules = new List<Module>();

            // Подсчет загруженных модулей
            int uploadModules = 0;
            string pathToModules = Path.Combine(_pahtToLoaderDll, @"Modules");
            DirectoryInfo directoryInfo = new DirectoryInfo(pathToModules);
            foreach(DirectoryInfo dir in directoryInfo.GetDirectories())
            {
                // Флаг на факт загрузки модуля
                bool isModuleLoad = false;
                string moduleName = dir.Name;
                string moduleVersion = string.Empty;

                String[] spearator = { ".dll" };
                
                // Активация библиотеки (без имплементации IExternalModule)
                if (moduleName.Contains("Library"))
                {
                    foreach (FileInfo file in dir.GetFiles())
                    {
                        if (CheckModuleName(file.Name, spearator))
                        {
                            // Каждую dll библиотеки - нужно прогрузить в текущее приложение, чтобы появилась связь в проекте.
                            // Если этого не сделать - будут проблемы с использованием загружаемых библиотек
                            System.Reflection.Assembly moduleAssembly = System.Reflection.Assembly.LoadFrom(file.FullName);
                            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                            if (moduleAssembly.FullName.Contains("KPLN"))
                                moduleVersion = fvi.FileVersion;

                            isModuleLoad = true;
                        }
                    }

                    if (isModuleLoad)
                    {
                        // Вывод в окно пользователя и лог
                        string msg = $"Модуль-библиотека [{moduleName}] версии {moduleVersion} успешно активирован!";
                        _logger.Info(msg);
                        uploadModules++;
                    }
                }
                // Ативация модуля (с имплементацией IExternalModule)
                else
                {
                    foreach (FileInfo file in dir.GetFiles())
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
                                    _moduleInstances.Add(moduleInstance);

                                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                                    moduleVersion = fvi.FileVersion;

                                    // Вывод в окно пользователя и лог
                                    string msg = $"Модуль [{moduleName}] версии {moduleVersion} успешно активирован!";
                                    _logger.Info(msg);
                                    uploadModules++;
                                    isModuleLoad = true;
                                }
                                else
                                {
                                    // Вывод в окно пользователя и лог
                                    string msg = $"Модуль [{moduleName}] не активирован.";
                                    _logger.Error(msg + $" Проблема: \n{loadingResult}");
                                }
                            }
                        }
                    }
                }

                if (!isModuleLoad)
                {
                    // Вывод в окно пользователя и лог
                    string msg = $"Модуль/библиотека {moduleName} - не активирован. Отсутсвует dll для активации";
                    _logger.Error(msg);
                }
            }

            return modules;
        }

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
                        MainInstruction = "Произошла серьезная ошибка при выполнении команды, обратись за помощью в BIM-отдел",
                        FooterText = $"См. файлы логов KPLN_Loader: {MainCashFolder}\\KPLN_Logs\\{RevitVersion}"
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
            
            if (CurrentRevitUser == null ) 
                return;
            
            if (CurrentRevitUser.IsExtraNet)
                _gtabService.SetRevitUserName(app.Username, CurrentRevitUser);
            else
                _dbService.SetRevitUserName(app.Username, CurrentRevitUser);
        }
    }
}
