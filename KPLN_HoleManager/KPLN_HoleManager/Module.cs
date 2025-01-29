using Autodesk.Revit.UI;
using KPLN_HoleManager.ExternalCommand;
using KPLN_Loader.Common;
using System.Reflection;

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
            // Создаём вкладку
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException) {}

            // Добавляем панель на вкладку
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Управление отверстиями");

            // Добавляем кнопку в панель
            AddPushButtonDataInPanel(
                "Менеджер\n отверстий",
                "Менеджер\n отверстий",
                "Открыть панель менеджера отверстий",
                "Через данную панель происходит обмен заданиями на отверстия между смежными отделами.",
                typeof(KPLN_HoleManager.Common.CommandShowDockablePane).FullName,
                panel,
                "KPLN_HoleManager.Resources.OpenManager.png",
                "http://moodle.stinproject.local"
            );

            // Регистрируем панель
            KPLN_HoleManager.Common.DockablePreferences.EnsureDockablePaneRegistered(application);

            // Подписываемся на событие открытия документа
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
          
            return Result.Succeeded;
        }

        // Обработчик события открытия документа
        private void OnDocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs args)
        {
            try
            {
                UIApplication uiApp = new UIApplication(sender as Autodesk.Revit.ApplicationServices.Application);
                KPLN_HoleManager.Common.DockablePreferences.HideDockablePane(uiApp);
            }
            catch
            {
                TaskDialog.Show("Ошибка", "При обработке плагина 'Менеджер отверстий' произошла ошибка. Перед использыванием плагина закройте его и откройте заново." +
                    "В случае, если ошибка повториться - обратитесь в BIM-отдел.");
            }
        }

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

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPathname)
        {
            var stream = GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(
                stream,
                System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat,
                System.Windows.Media.Imaging.BitmapCacheOption.Default
            );
            return decoder.Frames[0];
        }
    }
}