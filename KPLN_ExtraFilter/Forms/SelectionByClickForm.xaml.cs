using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByClickForm : Window
    {
#if !Debug2020 && !Revit2020
        private readonly ExternalEvent _externalEvent;
        private SelectionChangedHandler _handler;
#endif

        public SelectionByClickForm(UIApplication uiapp, Element userSelElem)
        {
            Document doc = uiapp.ActiveUIDocument.Document;
            CurrentSelectionByClickVM = new SelectionByClickVM(doc);

            InitializeComponent();

            DataContext = CurrentSelectionByClickVM;


            #region Блок настройки подписки на события окна на происходящие триггеры
#if !Debug2020 && !Revit2020
            _externalEvent = FormEventSubscriptionHelper.CreateSelectionChangedEvent(handler =>
            {
                _handler = handler;
                _handler.CurrentSelByClickVM = CurrentSelectionByClickVM;
            });

            ExternalEvent unsubExtEv = FormEventSubscriptionHelper.CreateSelectionUnsubscribeEvent(OnSelectionChanged);
#endif
            // Предустановка
            if (userSelElem != null)
            {
                this.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelDoc = doc;
                this.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelElem = userSelElem;
            }

#if !Debug2020 && !Revit2020
            FormEventSubscriptionHelper.SubscribeSelectionChanged(uiapp, this, OnSelectionChanged, unsubExtEv);
#endif
#endregion
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByClickVM CurrentSelectionByClickVM { get; set; }

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => _externalEvent?.Raise();
#endif

        /// <summary>
        /// Метод для фокуса на нужном CB. Триггер усложняет жизнь, для его обслуживания нужен доп. класс
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CHB_WhatParamData_Checked(object sender, RoutedEventArgs e) => this.CB_FilterParams.Focus();
    }
}
