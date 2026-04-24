using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SetParamsByFrameForm : Window
    {
        private ExternalEvent _viewExtEv;
        private ViewActivatedHandler _viewHandler;

        private ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;

        public SetParamsByFrameForm(UIApplication uiapp, IEnumerable<Element> userSelElems)
        {
            CurrentSetParamsByFrameVM = new SetParamsByFrameVM(this, uiapp.ActiveUIDocument.Document, userSelElems);

            InitializeComponent();

            DataContext = CurrentSetParamsByFrameVM;

            #region Блок настройки подписки на события окна на происходящие триггеры
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
            this.SetExternalEvent(viewExtEv, viewHandler, selExtEv, selHandler);
#endif

            // Подписываюсь на OnViewChanged
            uiapp.ViewActivated += OnViewChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            this.Closed += (s, e) => unsubViewExtEv.Raise();

#if !Debug2020 && !Revit2020
            // Подписываюсь на SelectionChanged
            uiapp.SelectionChanged += OnSelectionChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            this.Closed += (s, e) => unsubSelExtEv.Raise();
#endif
            #endregion
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SetParamsByFrameVM CurrentSetParamsByFrameVM { get; set; }

        public void SetExternalEvent(ExternalEvent viewExtEv, ViewActivatedHandler viewHandler, ExternalEvent selExtEv, SelectionChangedHandler selHandler)
        {
            _viewExtEv = viewExtEv;
            
            _viewHandler = viewHandler;
            _viewHandler.CurrentSetParamsByFrameVM = CurrentSetParamsByFrameVM;


            _selExtEv = selExtEv;
            
            _selHandler = selHandler;
            _selHandler.CurrentSetParamsByFrameVM = CurrentSetParamsByFrameVM;
        }

        public void RaiseUpdateSelChanged() => _selExtEv?.Raise();

        public void RaiseUpdateViewChanged() => _viewExtEv?.Raise();

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => this.RaiseUpdateViewChanged();

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => this.RaiseUpdateSelChanged();
#endif
    }
}
