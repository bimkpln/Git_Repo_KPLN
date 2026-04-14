using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.Services;
using KPLN_Library_PluginActivityWorker;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectionByModelExtCmd : IExternalCommand
    {
        /// <summary>
        /// Имя плагина. Использую в KPLN_DefaultPanelExtension_Modify
        /// </summary>
        public const string PluginName = "Дерево элементов";
        private SelectionByModel _mainForm;

        public Result ExecuteByUIApp(UIApplication uiapp, ViewFilterMode viewFilterMode)
        {
            //Получение объектов приложения и документа
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            _mainForm = new SelectionByModel(uiapp, viewFilterMode);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(_mainForm);
            
            
            // Создание ExternalEvent для переключения видов
            ViewActivatedHandler viewHandler = new ViewActivatedHandler();
            ExternalEvent viewExtEv = ExternalEvent.Create(viewHandler);

            // Создание ExternalEvent для отписки от переключения видов (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubViewActHandler unsubViewHandler = new UnsubViewActHandler() { Handler = OnViewChanged };
            ExternalEvent unsubViewExtEv = ExternalEvent.Create(unsubViewHandler);

#if !Debug2020 && !Revit2020
            // Создание ExternalEvent для выделения эл-в
            SelectionChangedHandler selHandler = new SelectionChangedHandler();
            ExternalEvent selExtEv = ExternalEvent.Create(selHandler);

            // Создание ExternalEvent для отписки от выбора эл-в (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubEventHandler unsubSelHandler = new UnsubEventHandler() { Handler = OnSelectionChanged };
            ExternalEvent unsubSelExtEv = ExternalEvent.Create(unsubSelHandler);
            
            // Доп настройки окна
            _mainForm.SetExternalEvent(viewExtEv, viewHandler, selExtEv, selHandler);
#endif

            // Подписываюсь на OnViewChanged
            uiapp.ViewActivated += OnViewChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubViewExtEv.Raise();

#if !Debug2020 && !Revit2020
            // Подписываюсь на SelectionChanged
            uiapp.SelectionChanged += OnSelectionChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubSelExtEv.Raise();
#endif


            _mainForm.Show();


            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            return Result.Succeeded;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) => ExecuteByUIApp(commandData.Application, ViewFilterMode.CurrentView);

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => _mainForm.RaiseUpdateViewChanged();

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => _mainForm.RaiseUpdateSelChanged();
#endif
    }
}
