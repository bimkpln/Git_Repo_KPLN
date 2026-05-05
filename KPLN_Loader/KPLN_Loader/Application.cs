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
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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
        /// Глобальная пометка лоудера, на работу в закрытом окружении KPLN
        /// </summary>
        public static bool IsExtraNet { get; private set; }

        /// <summary>
        /// Основной путь к основным конфигам для работы всей экосистемы
        /// </summary>
        public static string MainConfigPath { get; private set; }

        internal static LoaderStatusForm LoaderStatusForm { get; private set; }

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
        /// <summary>
        /// Коллекция зарегестрированных КАСТОМНЫХ кнопок в Autodesk.Windows
        /// </summary>
        public static List<(Autodesk.Windows.RibbonButton RButton, string IconBaseName, string KPLNPluginAssembleName)> KPLNWindButtonsForImageReverse = new List<(Autodesk.Windows.RibbonButton, string, string)>();

        /// <summary>
        /// Коллекция зарегестрированных кнопок для смены цвета
        /// </summary>
        public static List<(RibbonButton RButton, string IconBaseName, string KPLNPluginAssembleName)> KPLNButtonsForImageReverse = new List<(RibbonButton, string, string)>();

        /// <summary>
        /// Коллекция зарегестрированных кнопок В СТЭКЕ для смены цвета
        /// </summary>
        public static List<(Autodesk.Windows.RibbonItem RItem, string IconBaseName, string KPLNPluginAssembleName)> KPLNStackButtonsForImageReverse = new List<(Autodesk.Windows.RibbonItem, string, string)>();
#endif

        internal delegate void RiseStepProgress(MainStatus mainStatus, string toolTip, Brush brush);
        /// <summary>
        /// Событие, которое посылает сигналы в форму статуса загрузки
        /// </summary>
        internal event RiseStepProgress Progress;

        internal delegate void RiseLoadEvant(LoaderEvantEntity lModule, Brush brush);
        /// <summary>
        /// Событие, которое посылает сигналы в форму статуса загрузки
        /// </summary>
        internal event RiseLoadEvant LoadStatus;

        private readonly static string _diteTime = DateTime.Now.ToString("dd/MM/yyyy_HH/mm/ss");
        private static string _pahtToLoaderDll;
        private readonly List<IExternalModule> _moduleInstances = new List<IExternalModule>();
        private string _ribbonName;
        private SQLiteService _sqliteService;
        private MySQLService _mysqlService;
        private EnvironmentService _envService;
        private ModuleLoaderService _moduleLoader;
        private Logger _logger;

        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= OnIdling;
            application.ControlledApplication.DocumentOpened -= OnDocumentOpened;

            foreach(IExternalModule module in _moduleInstances)
            {
                module.Close();
            }
            
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {            
            // Инит основных полей/свойств
            InitializeRuntimeSettings(application);
            InitializeInfrastructure();
            string mainDbPath = GetMainDbPath();
            LoaderStatusForm = CreateStatusFormForLocalStartup(); 


            // Активация спец. окружения
            if (!InitializeEnvironment())
            {
                _logger.Info($"Неудачная попытка создания окружения в Revit {RevitVersion}\n");
                return Result.Failed;
            }


            // Активация модулей
            InitializeModuleLoader();
            Result startupResult = IsExtraNet
                ? InitializeExtraNetStartup(application, mainDbPath)
                : InitializeLocalStartup(application, mainDbPath);
            
            if (startupResult != Result.Succeeded)
            {
                _logger.Info($"Неудачная попытка инициализации в Revit {RevitVersion}\n");
                return startupResult;
            }

            
            // Фиксация успешного итога
            _logger.Info($"Успешная инициализация в Revit {RevitVersion}\n");
            SubscribeEvents(application);

            return Result.Succeeded;
        }

        /// <summary>
        /// Очистка от старых логов
        /// </summary>
        /// <param name="logPath">Путь к логам</param>
        public static void ClearingOldLogs(Logger logger, string logPath, string logName, bool writeClearingStatus = false)
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
                            if (writeClearingStatus)
                                logger.Error($"При попытке очистки старых логов произошла ошибка: {ex.Message}");
                        }
                    }
                }

                if (writeClearingStatus)
                    logger.Info($"Очистка от старых логов произведена успешно! Удалено {cnt}");
            }
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки в зависимости от темы
        /// </summary>
        /// <param name="iconBaseName">Чистое имя иконки</param>
        /// <param name="size">Размер иконки</param>
        public static ImageSource GetBtnImage_ByTheme(string assemblyName, string iconBaseName, int size)
        {
            string themeSuffix = string.Empty;

            UITheme theme = UIThemeManager.CurrentTheme;
            if (theme == UITheme.Dark)
                themeSuffix = "_dark";

            string fileName = $"{iconBaseName}{size}{themeSuffix}.png";
            string uriString = $"pack://application:,,,/{assemblyName};component/Imagens/{fileName}";

            ImageSource result = null;
            try
            {
                result = new BitmapImage(new Uri(uriString));
            }
            // Если нет дарк темы, то просто имя файла
            catch (IOException)
            {
                try
                {
                    fileName = $"{iconBaseName}{size}.png";
                    uriString = $"pack://application:,,,/{assemblyName};component/Imagens/{fileName}";
                    var a = new Uri(uriString).AbsolutePath;
                    var aa = new Uri(uriString).LocalPath;

                    result = new BitmapImage(new Uri(uriString));
                }
                catch
                {
                    // Если его нет - подсвечиваю ошибку разработчику
                    if (CurrentRevitUser != null && CurrentRevitUser.IsDebugMode)
                    {
                        MessageBox.Show($"Ошибка поиска картинки для плагина {assemblyName}. Имя картинки {fileName} " +
                                $"Разработчик - проверь структуру хранения данных, картинки должны быть в папке \"Imagens\", и должны быть ресурсом (Resource)",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }

            return result;
        }

        private void InitializeRuntimeSettings(UIControlledApplication application)
        {
            RevitVersion = application.ControlledApplication.VersionNumber;
            _pahtToLoaderDll = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

#if ExtraNet
            IsExtraNet = true;
            MainConfigPath = _pahtToLoaderDll + "\\KPLN_ExtraNet_Loader_Config.json";
            _ribbonName = "KPLN_ExtraNet";
#else
            IsExtraNet = false;
#if DEBUG
            MainConfigPath = _pahtToLoaderDll + "\\KPLN_Loader_Config.json";
#else
            MainConfigPath = _pahtToLoaderDll + "\\KPLN_Loader_Config.json";
#endif

            _ribbonName = "KPLN";
#endif
        }

        private void InitializeInfrastructure()
        {
            LogManager.Setup().LoadConfigurationFromFile(_pahtToLoaderDll + "\\nlog.config");
            _logger = LogManager.GetLogger("KPLN_Loader");

            string windrive = Path.GetPathRoot(Environment.SystemDirectory);
            MainCashFolder = $"{windrive}KPLN_Temp";

            string logDirPath = $"{MainCashFolder}\\KPLN_Logs\\{RevitVersion}";
            string logFileName = "KPLN_Loader";
            LogManager.Configuration.Variables["loader_logdir"] = logDirPath;
            LogManager.Configuration.Variables["loader_logfilename"] = logFileName;

            // Старт записи в лог
            _logger.Info($"Запуск в Revit {RevitVersion}. Версия модуля: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}. Статус IsExtraNet: {IsExtraNet}");

            Task.Run(() => ClearingOldLogs(_logger, logDirPath, logFileName, true));
        }

        private string GetMainDbPath() => EnvironmentService.DatabaseConfigs.FirstOrDefault(d => d.Name.Contains(EnvironmentService.DatabaseConfigs_LoaderMainDB)).Path;

        private bool InitializeEnvironment()
        {
            _envService = new EnvironmentService(_logger, RevitVersion, _diteTime, LoaderStatusForm);

            return _envService.ConfigFileChecker() && _envService.PreparingAndCliningDirectories();
        }

        private void InitializeModuleLoader()
        {
            _moduleLoader = new ModuleLoaderService(_logger, _ribbonName, _moduleInstances, (msg, isError) =>
            {
                LoadStatus?.Invoke(new LoaderEvantEntity(msg), isError ? Brushes.Red : Brushes.Black);
            });
        }

        private Result InitializeExtraNetStartup(UIControlledApplication application, string mainDbPath)
        {
            // Проверка подключения к БД
            try
            {
                _mysqlService = new MySQLService(_logger, mainDbPath);
                Progress?.Invoke(MainStatus.Envirnment, "Успешно!", Brushes.Green);
            }
            catch (Exception ex)
            {
                LoaderMessageService.ShowWarning(
                    $"{LoaderMessageService.UserMessages.DbConnectionFailedPrefix}\n" +
                    $"{LoaderMessageService.UserMessages.ContactSupport}\n\n" +
                    $"Текст ошибки: {ex.Message}");
                return Result.Cancelled;
            }

            // Инициализация
            try
            {
                CurrentRevitUser = _mysqlService.Authorization(_envService, IsExtraNet);
                if (CurrentRevitUser == null)
                {
                    LoaderStatusForm?.Dispatcher.Invoke(() => LoaderStatusForm.Close());
                
                    LoaderMessageService.ShowWarning(
                        $"{LoaderMessageService.UserMessages.InitializationFailed}\n" +
                        $"{LoaderMessageService.UserMessages.ContactSupport}");
                
                    return Result.Cancelled;
                }

                if (CurrentRevitUser.IsUserRestricted)
                {
                    LoaderStatusForm?.Dispatcher.Invoke(() => LoaderStatusForm.Close());

                    LoaderMessageService.ShowWarning(LoaderMessageService.UserMessages.RestrictedAccess);
                
                    return Result.Cancelled;
                }

                application.CreateRibbonTab(_ribbonName);
            
                LoaderStatusForm.SetStatuses(CurrentRevitUser.IsDebugMode, IsExtraNet);
                bool isUserDataUpdated = _mysqlService.SetUserLastConnectionDate(CurrentRevitUser);
                CurrentSubDepartment = _mysqlService.GetSubDepartmentForCurrentUser(CurrentRevitUser);

                LoaderDescription loaderDescription = _mysqlService.GetDescriptionForCurrentUser(CurrentRevitUser);
                LoaderStatusForm.SetInstruction(loaderDescription);
                LoaderStatusForm.LikeStatus += LoaderStatusForm_RiseLikeEvant;

                if (isUserDataUpdated && CurrentSubDepartment != null)
                    Progress?.Invoke(MainStatus.DbConnection, "Успешно!", Brushes.Green);
                else
                    Progress?.Invoke(MainStatus.DbConnection, "Замечания при подключения к БД! Подробнее - см. файл логов", Brushes.Orange);

                LoadStatus?.Invoke(
                        new LoaderEvantEntity($"Пользователь: [{CurrentRevitUser.Surname} {CurrentRevitUser.Name}], отдел [{CurrentSubDepartment.Code}]"),
                        Brushes.OrangeRed);

                _logger.Info("Активация модулей для пользователя:");

                string pathToModules = Path.Combine(_pahtToLoaderDll, @"Modules");
                DirectoryInfo modulesDirectory = new DirectoryInfo(pathToModules);
                _moduleLoader.LoadExtraNetModules(application, modulesDirectory);
                Progress?.Invoke(MainStatus.ModulesActivation, "Успешно!", Brushes.Green);
            }
            catch (Exception ex)
            {
                _logger.Error($"Глобальная ошибка плагина загрузки: \n{ex}");
                _logger.Info("Инициализация не удалась\n");
                LoaderStatusForm.Dispatcher.Invoke(() => LoaderStatusForm.Close());

                Exception currentEx = ex.InnerException ?? ex;
                LoaderMessageService.ShowWarning(string.Format(LoaderMessageService.UserMessages.GlobalInitializationError, currentEx.Message));
                return Result.Cancelled;
            }

            LoaderStatusForm.Start_WindowClose();
            return Result.Succeeded;
        }

        private Result InitializeLocalStartup(UIControlledApplication application, string mainDbPath)
        {
            try
            {
                Progress?.Invoke(MainStatus.Envirnment, "Успешно!", Brushes.Green);

                _sqliteService = new SQLiteService(_logger, mainDbPath);
                CurrentRevitUser = _sqliteService.Authorization(_envService, IsExtraNet);
                if (CurrentRevitUser == null)
                {
                    Progress?.Invoke(MainStatus.DbConnection, "Критическая ошибка пользователя! Подробнее - см. файл логов", Brushes.Red);
                    return Result.Cancelled;
                }

                application.CreateRibbonTab(_ribbonName);

                LoaderStatusForm.SetStatuses(CurrentRevitUser.IsDebugMode, IsExtraNet);
                bool isUserDataUpdated = _sqliteService.SetUserLastConnectionDate(CurrentRevitUser);
                CurrentSubDepartment = _sqliteService.GetSubDepartmentForCurrentUser(CurrentRevitUser);

                LoaderDescription loaderDescription = _sqliteService.GetDescriptionForCurrentUser(CurrentRevitUser);
                LoaderStatusForm.SetInstruction(loaderDescription);
                LoaderStatusForm.LikeStatus += LoaderStatusForm_RiseLikeEvant;

                if (isUserDataUpdated && CurrentSubDepartment != null)
                    Progress?.Invoke(MainStatus.DbConnection, "Успешно!", Brushes.Green);
                else
                    Progress?.Invoke(MainStatus.DbConnection, "Замечания при подключения к БД! Подробнее - см. файл логов", Brushes.Orange);

                LoadStatus?.Invoke(
                    new LoaderEvantEntity($"Пользователь: [{CurrentRevitUser.Surname} {CurrentRevitUser.Name}], отдел [{CurrentSubDepartment.Code}]"),
                    Brushes.OrangeRed);

                _logger.Info("Подготовка, копирование библиотек и активация модулей для пользователя:");
                IEnumerable<Module> userAllModules = _sqliteService.GetModulesForCurrentUser(CurrentRevitUser);
                int uploadModules = _moduleLoader.LoadLocalModules(application, userAllModules, m => _envService.CopyModule(m, CurrentRevitUser.IsDebugMode));

                if(userAllModules.Count() == 0)
                {
                    _logger.Error($"Модули отсутсвуют");
                    Progress?.Invoke(MainStatus.ModulesActivation, $"Модули отсутсвуют. Обратись к разработчику!", Brushes.Red);
                }
                else if (uploadModules == userAllModules.Count())
                    Progress?.Invoke(MainStatus.ModulesActivation, "Успешно!", Brushes.Green);
                else
                    Progress?.Invoke(MainStatus.ModulesActivation, $"С замечаниями. Смотри файл логов KPLN_Loader: {MainCashFolder}\\KPLN_Logs\\", Brushes.Orange);
            }
            catch (Exception ex)
            {
                _logger.Error($"Глобальная ошибка плагина загрузки: \n{ex}");
                _logger.Info("Инициализация не удалась\n");
                LoaderStatusForm.Dispatcher.Invoke(() => LoaderStatusForm.Close());

                Exception currentEx = ex.InnerException ?? ex;
                LoaderMessageService.ShowWarning(string.Format(LoaderMessageService.UserMessages.GlobalInitializationError, currentEx.Message));
                return Result.Cancelled;
            }

            LoaderStatusForm.Start_WindowClose();
            return Result.Succeeded;
        }

        /// <summary>
        /// Создание и запуск статус-окна для локального старта.
        /// </summary>
        private LoaderStatusForm CreateStatusFormForLocalStartup()
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

            return loaderStatusForm;
        }

        private void SubscribeEvents(UIControlledApplication application)
        {
            application.Idling += OnIdling;
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            application.ThemeChanged += OnThemeChanged;
#endif
        }

        /// <summary>
        /// Обработчик события RiseLikeEvant
        /// </summary>
        private void LoaderStatusForm_RiseLikeEvant(MainDB_LoaderDescriptions_RateType rateType, LoaderDescription loaderDescription) => _sqliteService.SetLoaderDescriptionUserRank(rateType, loaderDescription);

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
                _mysqlService.SetRevitUserName(app.Username, CurrentRevitUser);
            else
                _sqliteService.SetRevitUserName(app.Username, CurrentRevitUser);
        }

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
        /// <summary>
        /// Событие, происходящее при смене темы начиная с ревит 2024
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnThemeChanged(object sender, ThemeChangedEventArgs e)
        {
            foreach (var (rButton, iconBaseName, kplnPluginAssembleName) in KPLNButtonsForImageReverse)
            {
                rButton.Image = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 16);
                rButton.LargeImage = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 32);
            }

            foreach (var (rItem, iconBaseName, kplnPluginAssembleName) in KPLNStackButtonsForImageReverse)
            {
                rItem.Image = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 16);
                rItem.LargeImage = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 32);
            }

            foreach (var (rButton, iconBaseName, kplnPluginAssembleName) in KPLNWindButtonsForImageReverse)
            {
                rButton.Image = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 16);
                rButton.LargeImage = GetBtnImage_ByTheme(kplnPluginAssembleName, iconBaseName, 32);
            }
        }
#endif
    }
}
