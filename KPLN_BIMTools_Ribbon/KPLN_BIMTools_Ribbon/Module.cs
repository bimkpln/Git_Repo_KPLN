using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.ExternalCommands;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using NLog;
using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace KPLN_BIMTools_Ribbon
{
    public class Module : IExternalModule
    {
        internal static Logger CurrentLogger { get; private set; }

        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
        private UIApplication _uiApp;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            #region Получаю UIApplication из internal свойства UIControlledApplication
            // https://stackoverflow.com/questions/42382320/getting-the-current-application-and-document-from-iexternalapplication-revit
            string fieldName = "m_uiapplication";
            FieldInfo fi = application.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
            _uiApp = (UIApplication)fi.GetValue(application);
            #endregion

            #region Настройка NLog
            // Конфиг для логгера лежит в KPLN_Loader. Это связано с инициализацией dll самим ревитом. Настройку тоже производить в основном конфиге
            CurrentLogger = LogManager.GetLogger("KPLN_BIMTools");

            string windrive = $"{Path.GetPathRoot(Environment.SystemDirectory)}KPLN_Temp";
            string logDirPath = $"{windrive}\\KPLN_Logs\\{ModuleData.RevitVersion}";
            string logFileName = "KPLN_BIMTools";
            LogManager.Configuration.Variables["bimtools_logdir"] = logDirPath;
            LogManager.Configuration.Variables["bimtools_logfilename"] = logFileName;
            #endregion

            CommandRVTExchange.SetStaticEnvironment(application);

            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));

            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "BIM");

            #region Выпадающий список "Выгрузка"
            PulldownButtonData uploadPullDownData = new PulldownButtonData("Выгрузка", "Выгрузка")
            {
                ToolTip = "Плагины по выгрузке моделей",
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainLoad", 32),
            };
            PulldownButton uploadPullDown = panel.AddItem(uploadPullDownData) as PulldownButton;
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((uploadPullDown, "mainLoad", _assemblyName));
#endif

            //Добавляю кнопки в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                CommandAutoExchangeConfig.PluginName,
                CommandAutoExchangeConfig.PluginName,
                "Конфигурация автозапуска обмена Revit-моделями (плагины по обмену RVT и NW файлов)",
                string.Format(
                    "Пакетная выгрзка моделей.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandAutoExchangeConfig).FullName,
                uploadPullDown,
                "asConfig",
                "http://moodle/mod/book/view.php?id=502&chapterid=1300",
                true
            );

            AddPushButtonDataInPullDown(
                CommandRVTExchange.PluginName,
                CommandRVTExchange.PluginName,
                "Обмен Revit-моделями:\n" +
                "1. Экспорт/импорт моделей с Revit-Server KPLN на другой Revit-Server KPLN;\n" +
                "2. Экспорт/импорт моделей с Revit-Server KPLN на сервер KPLN и наоборот;\n" +
                "3. Копирование моделей внутри сервера KPLN.",
                string.Format(
                    "Пакетная выгрзка моделей.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandRVTExchange).FullName,
                uploadPullDown,
                "load",
                "http://moodle/mod/book/view.php?id=502&chapterid=1300",
                true
            );

            AddPushButtonDataInPullDown(
                CommandNWExport.PluginName,
                CommandNWExport.PluginName,
                "Экспорт моделей в Navisworks",
                string.Format(
                    "Пакетный (по предварительным настройкам) экспорт моделей в Navisworks.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandNWExport).FullName,
                uploadPullDown,
                "nwExport",
                "http://moodle/mod/book/view.php?id=502&chapterid=1300",
                true
            );
            #endregion

            #region Выпадающий список "Параметры"
            PulldownButtonData paramPullDownData = new PulldownButtonData("Параметры", "Параметры")
            {
                ToolTip = "Плагины по работе с параметрами",
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainParam", 32),
            };
            PulldownButton paramPullDown = panel.AddItem(paramPullDownData) as PulldownButton;
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((paramPullDown, "mainParam", Assembly.GetExecutingAssembly().GetName().Name));
#endif

            //Добавляю кнопки в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "LT: Экспорт",
                "LT: Экспорт",
                "Экспорт парамтеров из таблиц выбора (lookup tables), загруженных в семейство",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandLTablesExport).FullName,
                paramPullDown,
                "lookup",
                "http://moodle/",
                false
            );

            AddPushButtonDataInPullDown(
                "Пакетное добавление параметров",
                "Пакетное добавление параметров",
                "Пакетно добавить в семейство параметры",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(batchAddingParameters).FullName,
                paramPullDown,
                "batchAddingParameters",
                "http://moodle/mod/book/view.php?id=502&chapterid=1329",
                false
            );
            #endregion

            #region Менеджер работы с БД
            AddPushButtonDataInPanel(
                "Менеджер\nБД",
                "Менеджер\nБД",
                "Запускает менеджер работы с БД",
                string.Format(
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandDBManager).FullName,
                panel,
                "db",
                "http://moodle.stinproject.local",
                true
            );
            #endregion

            #region Выполняю автоматический запуск
            //Файл - флаг.Его наличие должно сигнализировать о автоматическом старте.
            string[] args = Environment.GetCommandLineArgs();
            foreach (string arg in args)
            {
                if (arg.Equals("AutoStart"))
                {
                    ExchangeService.IsAutoStart = true;
                    AutoRun(nameof(CommandRVTExchange), RevitDocExchangeEnum.Revit);
                }
            }
            #endregion

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для добавления кнопки в выпадающий список
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="pullDownButton">Выпадающий список, в который добавляем кнопку</param>
        /// <param name="imageName">Имя иконки, как ресурса</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPullDown(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            PulldownButton pullDownButton,
            string imageName,
            string contextualHelp,
            bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = pullDownButton.AddPushButton(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

            if (avclass) button.AvailabilityClassName = typeof(StaticAvailable).FullName;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }

        /// <summary>
        /// Метод для добавления отдельной кнопки в панель
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="panel">Панель, в которую добавляем кнопку</param>
        /// <param name="imageName">Имя иконки, как ресурса</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPanel(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            RibbonPanel panel,
            string imageName,
            string contextualHelp,
            bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

            if (avclass) button.AvailabilityClassName = typeof(StaticAvailable).FullName;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }

        /// <summary>
        /// Очистка от старых логов
        /// </summary>
        /// <param name="logPath">Путь к логам</param>
        private void ClearingOldLogs(string logPath, string logName)
        {
            if (Directory.Exists(logPath))
            {
                DirectoryInfo logDI = new DirectoryInfo(logPath);
                foreach (FileInfo log in logDI.EnumerateFiles())
                {
                    if (log.CreationTime.Date < DateTime.Now.AddDays(-5) && log.Name.Contains(logName))
                    {
                        try
                        {
                            log.Delete();
                        }
                        // Ошибка будет только если файл занят
                        catch (UnauthorizedAccessException) { }
                        catch (Exception ex)
                        {
                            CurrentLogger.Error($"При попытке очистки старых логов произошла ошибка: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void AutoRun(string externalCommand, RevitDocExchangeEnum exchangeEnum)
        {
            // Создаем тип
            Type type = Type.GetType($"KPLN_BIMTools_Ribbon.ExternalCommands.{externalCommand}", true);

            // Создаем экземпляр типа
            object instance = Activator.CreateInstance(type);

            // Определяем метод ExecuteByUIApp
            MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");

            // Вызываем метод ExecuteByUIApp, передавая _uiApp как аргумент
            if (executeMethod != null)
                executeMethod.Invoke(instance, new object[] { _uiApp, exchangeEnum });
            else
                throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");
        }
    }
}
