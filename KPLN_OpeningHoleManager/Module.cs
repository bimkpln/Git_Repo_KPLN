using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
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
        // Новый постоянный ID применяет вкладку диспетчера проекта вместо сохранённого отдельного размещения.
        internal static readonly DockablePaneId PaneId = new DockablePaneId(new Guid("943e18af-00a1-4c1b-a046-b23cb5055807"));
        internal static MainMenu MainMenuViewer;

        internal static UIControlledApplication CurrentUIContrApp { get; private set; }

        internal static UIApplication CurrentUIApplication { get; private set; }
        internal static string CurrentFileName { get; private set; }
        internal static DBProject CurrentDBProject { get; private set; }
        internal static Document CurrentDoc { get; private set; }
        internal static DBSubDepartment CurrnetDocSubDep { get; private set; }

        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private DateTime _nextSelectionCheck = DateTime.MinValue;
        private static Module _instance;
        private bool _tracking;

        public Result Close()
        {
            SetTracking(false);
            CurrentUIContrApp.DockableFrameVisibilityChanged -= Application_DockableFrameVisibilityChanged;
            _instance = null;

            return Result.Succeeded;
        }


        public Result Execute(UIControlledApplication application, string tabName)
        {
            _instance = this;
            CurrentUIContrApp = application;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

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

            // Постоянно слушаем только открытие/закрытие нашей панели.
            CurrentUIContrApp.DockableFrameVisibilityChanged += Application_DockableFrameVisibilityChanged;

            return Result.Succeeded;
        }

        internal static void OpenManager(UIApplication application)
        {
            _instance.SetTracking(true);
            SetActiveDocument(application, application.ActiveUIDocument?.Document);
            MainMenuViewer.MainMenu_VM.History.RequestRefresh(CurrentDoc, true);
        }

        private void Application_DockableFrameVisibilityChanged(object sender, DockableFrameVisibilityChangedEventArgs args)
        {
            if (!PaneId.Equals(args.PaneId))
                return;

            bool opening = args.DockableFrameShown && !_tracking;
            SetTracking(args.DockableFrameShown);

            if (!opening)
                return;

            UIApplication application = sender as UIApplication ?? CurrentUIApplication;
            if (application != null)
                OpenManager(application);
        }

        private void SetTracking(bool enabled)
        {
            if (_tracking == enabled)
                return;

            _tracking = enabled;
            _nextSelectionCheck = DateTime.MinValue;

            if (enabled)
            {
                MainMenuViewer.MainMenu_VM.History.Resume();

                CurrentUIContrApp.ViewActivated += Application_ViewActivated;
                CurrentUIContrApp.Idling += Application_Idling;
                CurrentUIContrApp.ControlledApplication.DocumentChanged += Application_DocumentChanged;
            }
            else
            {
                CurrentUIContrApp.ViewActivated -= Application_ViewActivated;
                CurrentUIContrApp.Idling -= Application_Idling;
                CurrentUIContrApp.ControlledApplication.DocumentChanged -= Application_DocumentChanged;

                // Отменяем уже поставленные в очередь действия и незаписанный автокэш.
                MainMenuViewer.MainMenu_VM.History.Suspend();
            }
        }


        private void Application_Idling(object sender, IdlingEventArgs args)
        {
            var application = sender as UIApplication;
            if (!_tracking || application == null)
                return;

            // После закрытия последней модели ViewActivated может не прийти.
            if (CurrentDoc != null && !CurrentDoc.IsValidObject)
                SetActiveDocument(application, null);

            var history = MainMenuViewer?.MainMenu_VM.History;
            if (history == null || !history.OnlySelected || DateTime.UtcNow < _nextSelectionCheck)
                return;

            _nextSelectionCheck = DateTime.UtcNow.AddMilliseconds(200);
            history.RefreshSelection(application.ActiveUIDocument);
        }

        private void Application_ViewActivated(object sender, ViewActivatedEventArgs args)
        {
            if (!_tracking || args.Document == null || !args.Document.IsValidObject)
                return;

            bool changed = CurrentDoc == null || !CurrentDoc.IsValidObject || !args.Document.Equals(CurrentDoc);
            SetActiveDocument(new UIApplication(args.Document.Application), args.Document);

            var history = MainMenuViewer.MainMenu_VM.History;
            if (changed)
                history.RequestRefresh(CurrentDoc, true);
            else if (history.OnlyActiveView)
                history.RequestVisibilityRefresh();
        }


        internal static void SetActiveDocument(UIApplication application, Document document)
        {
            CurrentUIApplication = application;
            if (document == null || !document.IsValidObject || document.IsFamilyDocument)
            {
                CurrentDoc = null;
                CurrentFileName = null;
                CurrentDBProject = null;
                CurrnetDocSubDep = null;
                MainMenuViewer.MainMenu_VM.History.SetDocument(null);
                return;
            }

            // Equals обращается к native-объектам обоих документов, включая аргумент.
            if (CurrentDoc != null && CurrentDoc.IsValidObject && document.Equals(CurrentDoc))
                return;

            CurrentDoc = document;
            MainMenuViewer.MainMenu_VM.History.SetDocument(document);
            CurrentFileName = document.IsWorkshared
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(document.GetWorksharingCentralModelPath())
                : document.PathName;
            try
            {
                CurrentDBProject = SQLiteMainService.SQLitePrjServiceInst.GetDBProject_ByRevitDocFileNameANDRVersion(CurrentFileName, ModuleData.RevitVersion);
            }
            catch
            {
                CurrentDBProject = null;
            }


            MainMenuViewer.MainMenu_VM.SetDataGeomParamsData(CurrentDBProject);
        }

        private void Application_DocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            // На закрытой панели обработчик отписан; флаг дополнительно защищает переходное состояние.
            if (!_tracking)
                return;

            Document document = args.GetDocument();
            if (document == null || !document.IsValidObject || CurrentDoc == null || !CurrentDoc.IsValidObject
                || !document.Equals(CurrentDoc))
                return;


            var transactionNames = args.GetTransactionNames();
            if (args.Operation == UndoOperation.TransactionCommitted && transactionNames.Count > 0
                && transactionNames.All(name => name == OpeningHistoryService.CacheTransactionName))
                return;

            // Здесь только ID и отметка создания. Чтение/запись истории выполняется позже из Idling.
            MainMenuViewer.MainMenu_VM.History.HandleDocumentChanged(document, args);
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
