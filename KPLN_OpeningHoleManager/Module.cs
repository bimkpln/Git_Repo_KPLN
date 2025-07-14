using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.ExternalCommands;
using KPLN_OpeningHoleManager.Forms;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_OpeningHoleManager
{
    public class Module : IExternalModule
    {
        internal static readonly DockablePaneId PaneId = new DockablePaneId(new Guid("71045bec-5187-420e-9a7e-93953bf2930c"));
        internal static MainMenu MainMenuViewer;

        internal static UIControlledApplication CurrentUIContrApp { get; private set; }

        internal static UIApplication CurrentUIApplication { get; private set; }
        internal static string CurrentFileName { get; private set; }
        internal static DBProject CurrentDBProject { get; private set; }
        internal static Document CurrentDoc { get; private set; }
        internal static DBSubDepartment CurrnetDocSubDep { get; private set; }
        internal static int RevitVersion { get; private set; }

        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            CurrentUIContrApp = application;
            RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            //Ищу или создаю панель инструменты
            string panelName = "Междисциплинарный анализ";
            RibbonPanel panel = null;
            IEnumerable<RibbonPanel> tryPanels = CurrentUIContrApp.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryPanels.Any())
                panel = tryPanels.FirstOrDefault();
            else
                panel = CurrentUIContrApp.CreateRibbonPanel(tabName, panelName);

            //Добавляю кнопку в панель (тут приведен пример поиска панели, вместо этого - панель можно создать)
            AddPushButtonDataInPanel(
                "Менеджер\nотверстий",
                "Менеджер\nотверстий",
                "Запускает отдельное меню менеджера отверстий",
                string.Format(
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ShowMainMenu).FullName,
                panel,
                "KPLN_OpeningHoleManager.Imagens.mainMenu_Big.png",
                "http://moodle.stinproject.local"
            );

            MainMenuViewer = new MainMenu();
            CurrentUIContrApp.RegisterDockablePane(PaneId, "Менеджер отверстий", MainMenuViewer);

            CurrentUIContrApp.ViewActivated += Application_ViewActivated;

            return Result.Succeeded;
        }

        private void Application_ViewActivated(object sender, ViewActivatedEventArgs args)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (args.Document == null || args.Document.IsFamilyDocument)
                return;

            string openViewFileName = args.Document.IsWorkshared
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(args.Document.GetWorksharingCentralModelPath())
                : args.Document.PathName;

            if (openViewFileName == CurrentFileName)
                return;

            CurrentUIApplication = new UIApplication(args.Document.Application);

            CurrentFileName = openViewFileName;
            DBProject openViewDBProject = MainDBService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(CurrentFileName, RevitVersion);
            if (openViewDBProject == null)
                return;

            CurrentDBProject = openViewDBProject;

            MainMenuViewer.MainMenu_VM.SetDataGeomParamsData(CurrentDBProject);
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
