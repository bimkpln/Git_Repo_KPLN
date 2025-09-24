using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Batch.Availability;
using NLog;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ModelChecker_Batch
{
    public class Module : IExternalModule
    {
        internal static Logger CurrentLogger { get; private set; }
        
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);
            ModuleData.RevitUIControlledApp = application;


            #region Настройка NLog
            // Конфиг для логгера лежит в KPLN_Loader. Это связано с инициализацией dll самим ревитом. Настройку тоже производить в основном конфиге
            CurrentLogger = LogManager.GetLogger("KPLN_ModelChecker_Batch");

            string windrive = $"{Path.GetPathRoot(Environment.SystemDirectory)}KPLN_Temp";
            string logDirPath = $"{windrive}\\KPLN_Logs\\{ModuleData.RevitVersion}";
            string logFileName = "KPLN_ModelChecker_Batch";
            LogManager.Configuration.Variables["modelCheckerBatch_logdir"] = logDirPath;
            LogManager.Configuration.Variables["modelCheckerBatch_logfilename"] = logFileName;
            #endregion

            Task clearingLogs = Task.Run(() => ClearingOldLogs(logDirPath, logFileName));


            //Добавляю кнопку в панель
            string currentPanelName = "Контроль качества";
            RibbonPanel currentPanel = application.GetRibbonPanels(tabName).Where(i => i.Name == currentPanelName).ToList().FirstOrDefault()
                ?? application.CreateRibbonPanel(tabName, "Контроль качества");

            //Добавляю кнопку в панель (тут приведен пример поиска панели, вместо этого - панель можно создать)
            AddPushButtonDataInPanel(
                "Пакетная\nпроверка",
                "Пакетная\nпроверка",
                "Пакетный запуск проверок моделей в автоматическом режиме",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandShowManager).FullName,
                currentPanel,
                "KPLN_ModelChecker_Batch.Imagens.mainIcon.png",
                "http://moodle/",
                true
            );

            return Result.Succeeded;
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
            button.Image = PngImageSource(imageName);
            button.LargeImage = PngImageSource(imageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

            if (avclass) button.AvailabilityClassName = typeof(StaticAvailable).FullName;
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
        private void ClearingOldLogs(string logPath, string logName)
        {
            if (Directory.Exists(logPath))
            {
                System.IO.DirectoryInfo logDI = new System.IO.DirectoryInfo(logPath);
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
    }
}
