using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.ExternalCommands;
using KPLN_Loader.Common;
using NLog;
using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_BIMTools_Ribbon
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private Logger _logger;
        private string _revitVersion;
        private UIApplication _uiApp;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            _revitVersion = application.ControlledApplication.VersionNumber;

            #region Получаю UIApplication из internal свойства UIControlledApplication
            // https://stackoverflow.com/questions/42382320/getting-the-current-application-and-document-from-iexternalapplication-revit
            string fieldName = "m_uiapplication";
            FieldInfo fi = application.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
            _uiApp = (UIApplication)fi.GetValue(application);
            #endregion

            #region Настройка NLog
            // Конфиг для логгера лежит в KPLN_Loader. Это связано с инициализацией dll самим ревитом. Настройку тоже производить в основном конфиге
            _logger = LogManager.GetLogger("KPLN_BIMTools");

            string logDirPath = $"c:\\temp\\KPLN_Logs\\{_revitVersion}";
            string logFileName = "KPLN_BIMTools";
            LogManager.Configuration.Variables["bimtools_logdir"] = logDirPath;
            LogManager.Configuration.Variables["bimtools_logfilename"] = logFileName;
            #endregion

            CommandRSExchange.SetStaticEnvironment(application, _logger, _revitVersion);

            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));

            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "BIM");

            #region Выпадающий список "Выгрузка"
            PulldownButtonData uploadPullDownData = new PulldownButtonData("Выгрузка", "Выгрузка")
            {
                ToolTip = "Плагины по выгрузке моделей",
                LargeImage = PngImageSource("KPLN_BIMTools_Ribbon.Imagens.mainLoadBig.png"),
            };
            PulldownButton uploadPullDown = panel.AddItem(uploadPullDownData) as PulldownButton;

            //Добавляю кнопки в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "RS: Обмен",
                "RS: Обмен",
                "Обмен (экспорт/импорт) моделей с Revit-Server KPLN",
                string.Format(
                    "Пакетная выгрзка моделей.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandRSExchange).FullName,
                uploadPullDown,
                "KPLN_BIMTools_Ribbon.Imagens.loadSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1300",
                true
            );

            AddPushButtonDataInPullDown(
                "NW: Экспорт",
                "NW: Экспорт",
                "Экспорт моделей в Navisworks",
                string.Format(
                    "Пакетный (по предварительным настройкам) экспорт моделей в Navisworks.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandNWExport).FullName,
                uploadPullDown,
                "KPLN_BIMTools_Ribbon.Imagens.nwExportSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1300",
                true
            );
            #endregion

            #region Выпадающий список "Параметры"
            PulldownButtonData paramPullDownData = new PulldownButtonData("Параметры", "Параметры")
            {
                ToolTip = "Плагины по работе с параметрами",
                LargeImage = PngImageSource("KPLN_BIMTools_Ribbon.Imagens.mainParamBig.png"),
            };
            PulldownButton paramPullDown = panel.AddItem(paramPullDownData) as PulldownButton;

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
                "KPLN_BIMTools_Ribbon.Imagens.lookupSmall.png",
                "http://moodle/",
                false
            );

            AddPushButtonDataInPullDown(
                "Пакетное добавление параметров",
                "Пакетное добавление параметров",
                "Нужно написать описание",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandLTablesExport).FullName,
                paramPullDown,
                "KPLN_BIMTools_Ribbon.Imagens.lookupSmall.png",
                "http://moodle/",
                false
            );
            #endregion

            #region Выполняю автоматический запуск
            //AutoRun(nameof(CommandRSExchange));
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
        private void AddPushButtonDataInPullDown(string name, string text, string shortDescription, string longDescription, string className, PulldownButton pullDownButton, string imageName, string contextualHelp, bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = pullDownButton.AddPushButton(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = PngImageSource(imageName);
            button.LargeImage = PngImageSource(imageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

            if (avclass)
            {
                button.AvailabilityClassName = typeof(StaticAvailable).FullName;
            }
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="embeddedPathname">Имя иконки с раширением</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
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
                            _logger.Error($"При попытке очистки старых логов произошла ошибка: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void AutoRun(string externalCommand)
        {
            // Создаем тип
            Type type = Type.GetType($"KPLN_BIMTools_Ribbon.ExternalCommands.{externalCommand}", true);

            // Создаем экземпляр типа
            object instance = Activator.CreateInstance(type);

            // Определяем метод ExecuteByUIApp
            MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");

            // Вызываем метод ExecuteByUIApp, передавая _uiApp как аргумент
            if (executeMethod != null)
                executeMethod.Invoke(instance, new object[] { _uiApp });
            else
                throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");
        }
    }
}
