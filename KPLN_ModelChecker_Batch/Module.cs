using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Batch.Availability;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ModelChecker_Batch
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);
            ModuleData.RevitUIControlledApp = application;

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
    }
}
