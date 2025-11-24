using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByClickForm : Window
    {
        private ExternalEvent _externalEvent;
        private SelectionChangedHandler _handler;

        public SelectionByClickForm(Document doc)
        {
            CurrentSelectionByClickVM = new SelectionByClickVM(doc);

            InitializeComponent();

            DataContext = CurrentSelectionByClickVM;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByClickVM CurrentSelectionByClickVM { get; set; }

        public void SetExternalEvent(ExternalEvent externalEvent, SelectionChangedHandler handler)
        {
            _externalEvent = externalEvent;
            _handler = handler;
            _handler.CurrentSelByClickVM = CurrentSelectionByClickVM;
        }

        public void RaiseUpdate() => _externalEvent?.Raise();

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                CurrentSelectionByClickVM.RunSelection();
        }

        private void CHB_WhatParamData_Checked(object sender, RoutedEventArgs e) => this.CB_FilterParams.Focus();
    }
}
