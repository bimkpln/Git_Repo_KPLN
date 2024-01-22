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

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            _revitVersion = application.ControlledApplication.VersionNumber;

            #region Настройка NLog
            LogManager.Setup().LoadConfigurationFromFile(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString() + "\\nlog.config");
            _logger = LogManager.GetLogger("KPLN_BIMTools");

            string logDirPath = $"c:\\temp\\KPLN_Logs\\{_revitVersion}";
            string logFileName = "KPLN_BIMTools";
            LogManager.Configuration.Variables["logdir"] = logDirPath;
            LogManager.Configuration.Variables["logfilename"] = logFileName;
            #endregion

            CommandExportToRS commandExportToRS = new CommandExportToRS(application, _logger);
            CommandImportFromRS commandImportFromRS = new CommandImportFromRS(application, _logger);

            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));

            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "BIM");

            //Добавляю выпадающий список pullDown
            PulldownButtonData pullDownData = new PulldownButtonData("Выгрзка", "Выгрзка")
            {
                ToolTip = "Плагины по выгрузке моделей",
                LargeImage = PngImageSource("KPLN_BIMTools_Ribbon.Imagens.mainLoadBig.png"),
            };
            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;

            //Добавляю кнопку в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "RS: Экспорт",
                "RS: Экспорт",
                "Экспортировать модели на Revit-Server KPLN",
                string.Format(
                    "Пакетная выгрзка моделей.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandExportToRS).FullName,
                pullDown,
                "KPLN_BIMTools_Ribbon.Imagens.loadSmall.png",
                "http://moodle.stinproject.local",
                true
            );

            //Добавляю кнопку в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "RS: Импорт",
                "RS: Импорт",
                "Импортировать модели с Revit-Server KPLN",
                string.Format(
                    "Пакетный импорт моделей.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandImportFromRS).FullName,
                pullDown,
                "KPLN_BIMTools_Ribbon.Imagens.loadSmall.png",
                "http://moodle.stinproject.local",
                true
            );

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
    }
}
