using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_CommandsWheel
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            Services.HotkeyService.Shutdown();
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            // Установка основных полей модуля
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);
            Services.HotkeyService.Initialize();

            //Добавляю панель
            RibbonPanel wheelsCommandsPanel = application.CreateRibbonPanel(tabName, "Штурвал команд");

            AddPushButtonDataInPanel(
                "KPLNCommandsWheelRun",
                "Штурвал",
                "Открыть штурвал команд",
                string.Format(
                    "Открыть штурвал команд. Эту команду можно назначить на горячую клавишу в Revit.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandsWheel).FullName,
                wheelsCommandsPanel,
                "KPLN_CommandsWheel.Imagens.commandsWheels.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1359"
            );

            AddPushButtonDataInPanel(
                "KPLNCommandsWheelSearch",
                "Команды",
                "Поиск команд и настройки штурвала",
                string.Format(
                    "Поиск команд на ленте Revit, избранное, последние команды, добавление в штурвал. Эту команду можно назначить на горячую клавишу в Revit.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandSearch).FullName,
                wheelsCommandsPanel,
                "KPLN_CommandsWheel.Imagens.settingsWheels.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1359"
            );

            Services.HotkeyService.Initialize();

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
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = PngImageSource(imageName);
            button.LargeImage = PngImageSource(imageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
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