using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_TaskManager.Common;
using KPLN_TaskManager.ExternalCommands;
using KPLN_TaskManager.Forms;
using KPLN_TaskManager.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_TaskManager
{
    public class Module : IExternalModule
    {
        internal static readonly DockablePaneId PaneId = new DockablePaneId(new Guid("B1234567-89AB-CDEF-0123-456789ABCDEF"));
        internal static TaskManagerView MainMenuViewer;

        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private static string _currentFileName;
        private static UIControlledApplication _uiContrApp;

        internal static DBProject CurrentDBProject { get; private set; }
        internal static UIApplication CurrentUIApplication { get; private set; }

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            _uiContrApp = application;

            //Ищу или создаю панель инструменты
            string panelName = "Инструменты";
            RibbonPanel panel = null;
            IEnumerable<RibbonPanel> tryPanels = _uiContrApp.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryPanels.Any())
                panel = tryPanels.FirstOrDefault();
            else
                panel = _uiContrApp.CreateRibbonPanel(tabName, panelName);

            //Добавляю кнопку в панель (тут приведен пример поиска панели, вместо этого - панель можно создать)
            AddPushButtonDataInPanel(
                "Менеджер\nзадач",
                "Менеджер\nзадач",
                "Запускает отдельное меню менеджера задач",
                string.Format(
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ShowMainMenu).FullName,
                panel,
                "KPLN_TaskManager.Imagens.showMenu.png",
                "http://moodle.stinproject.local"
            );

            MainMenuViewer = new TaskManagerView();
            _uiContrApp.RegisterDockablePane(PaneId, "Менеджер задач", MainMenuViewer);

            _uiContrApp.ControlledApplication.DocumentOpened += Application_DocumentOpened;
            _uiContrApp.ViewActivated += Application_ViewActivated;

            return Result.Succeeded;
        }

        private void Application_DocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (doc == null || doc.IsFamilyDocument)
                return;

            if (CurrentUIApplication == null)
                CurrentUIApplication = new UIApplication(doc.Application);

            _currentFileName = doc.IsWorkshared
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

            CurrentDBProject = MainDBService.ProjectDbService.GetDBProject_ByRevitDocFileName(_currentFileName);
            if (CurrentDBProject == null)
                return;

            IEnumerable<TaskItemEntity> docTasks = TaskManagerDBService.GetEntities_ByDBProject(CurrentDBProject);

            if (docTasks
                .Where(task => 
                    task.CreatedTaskDepartmentId == MainDBService.CurrentDBUserSubDepartment.Id 
                    || task.DelegatedDepartmentId == MainDBService.CurrentDBUserSubDepartment.Id)
                .Any(task =>
                    task.TaskStatus == TaskStatusEnum.Open))
                MainMenuViewer.LoadTaskData();
            
            ShowMainMenu.ShowPanel(_uiContrApp, false);
        }

        private void Application_ViewActivated(object sender, ViewActivatedEventArgs args)
        {
            Document doc = args.Document;
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (doc == null || doc.IsFamilyDocument)
                return;

            string openViewFileName = doc.IsWorkshared
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

            if (openViewFileName == _currentFileName)
                return;

            _currentFileName = openViewFileName;

            DBProject openViewDBProject = MainDBService.ProjectDbService.GetDBProject_ByRevitDocFileName(_currentFileName);
            if (openViewDBProject == null || CurrentDBProject == null || openViewDBProject.Id == CurrentDBProject.Id)
                return;

            CurrentDBProject = openViewDBProject;


            MainMenuViewer.LoadTaskData();
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
