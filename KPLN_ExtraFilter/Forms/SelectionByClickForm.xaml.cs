using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using KPLN_Library_Forms.Services;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByClickForm : Window
    {
        private ExternalEvent _externalEvent;
        private SelectionChangedHandler _handler;

        public SelectionByClickForm(UIApplication uiapp, Element userSelElem)
        {
            Document doc = uiapp.ActiveUIDocument.Document;
            CurrentSelectionByClickVM = new SelectionByClickVM(doc);

            InitializeComponent();

            DataContext = CurrentSelectionByClickVM;


#region Блок настройки подписки на события окна на происходящие триггеры
#if !Debug2020 && !Revit2020
            // Создание ExternalEvent для выделения эл-в
            SelectionChangedHandler selHandler = new SelectionChangedHandler();
            ExternalEvent selExtEv = ExternalEvent.Create(selHandler);

            // Создание ExternalEvent для отписки от выбора эл-в (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubEventHandler unsubHandler = new UnsubEventHandler() { Handler = OnSelectionChanged };
            ExternalEvent unsubExtEv = ExternalEvent.Create(unsubHandler);

            this.SetExternalEvent(selExtEv, selHandler);
#endif
            // Предустановка
            if (userSelElem != null)
            {
                this.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelDoc = doc;
                this.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelElem = userSelElem;
            }

#if !Debug2020 && !Revit2020
            // Подписываюсь на SelectionChanged
            uiapp.SelectionChanged += OnSelectionChanged;
            
            // Подписываю окно на отписку (через ExternalEvent)
            this.Closed += (s, e) => unsubExtEv.Raise();
#endif
#endregion
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByClickVM CurrentSelectionByClickVM { get; set; }

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => this.RaiseUpdate();
#endif

        public void SetExternalEvent(ExternalEvent externalEvent, SelectionChangedHandler handler)
        {
            _externalEvent = externalEvent;
            
            _handler = handler;
            _handler.CurrentSelByClickVM = CurrentSelectionByClickVM;
        }

        public void RaiseUpdate() => _externalEvent?.Raise();

        /// <summary>
        /// Метод для фокуса на нужном CB. Триггер усложняет жизнь, для его обслуживания нужен доп. класс
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CHB_WhatParamData_Checked(object sender, RoutedEventArgs e) => this.CB_FilterParams.Focus();
    }
}
