using Autodesk.Revit.UI;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using KPLN_Loader.Common;

namespace KPLN_HoleManager
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            // Регистрация DockablePane
            Common.DockablePreferences.RegisterDockablePane(application);

            // Добавление RibbonPanel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Менеджер отверстий");

            // Добавляю кнопку в панель
            AddPushButtonDataInPanel(
                "Открыть\nменеджер",
                "Открыть\nменеджер",
                "Открыть панель менеджера отверстий",
                string.Format(
                    "Через данную панель происходит обмен заданиями на отверстия между смежными отделами.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(Common.CommandShowDockablePane).FullName,
                panel,
                "KPLN_HoleManager.Imagens.OpenManager.png",
                "http://moodle.stinproject.local"
            );
            return Result.Succeeded;
        }

        // Метод для добавления отдельной в панель
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

        // Метод для добавления иконки для кнопки
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}