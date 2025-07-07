﻿using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.ExternalCommands;

namespace KPLN_ModelChecker_User.Forms
{

    public partial class CheckMonolithInfo : Window
    {
        private readonly UIDocument _uiDoc;

#if (Revit2023 || Debug2023)
        private readonly DeleteClashPointHandler _delHandler;
        private readonly ExternalEvent _delEvent;

        private readonly ShowClashPointHandler _showHandler;
        private readonly ExternalEvent _showEvent;
#endif

        public ObservableCollection<FamilyInstance> MonolithClashPoints { get; }
        public ObservableCollection<SkippedElementInfo> SkippedElements { get; }

        public CheckMonolithInfo(UIDocument uiDoc,
                         IList<FamilyInstance> clashPoints,
                         IList<(Element elem, string origin)> skipped)
        {
            InitializeComponent();
            _uiDoc = uiDoc;

            MonolithClashPoints = new ObservableCollection<FamilyInstance>(clashPoints);
            SkippedElements = skipped != null ? new ObservableCollection<SkippedElementInfo>( skipped.Select(pair => new SkippedElementInfo(pair.elem, pair.origin)))
                    : new ObservableCollection<SkippedElementInfo>();
            DataContext = this;
#if (Revit2023 || Debug2023)
            _delHandler = new DeleteClashPointHandler();
            _delEvent = ExternalEvent.Create(_delHandler);
            _showHandler = new ShowClashPointHandler();
            _showEvent = ExternalEvent.Create(_showHandler);
#endif
        }

        // ClashPoint. Кнопка «Выбрать элемент»
        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
#if (Revit2023 || Debug2023)
            if (sender is Button b && b.DataContext is FamilyInstance fi)
                _uiDoc.Selection.SetElementIds(new[] { fi.Id });
#endif
        }

        // ClashPoint. Кнопка «Создание обзорного 3D-вида»
        private void ShowButton_Click(object sender, RoutedEventArgs e)
        {
#if (Revit2023 || Debug2023)
            if (!(sender is Button b && b.DataContext is FamilyInstance fi)) return;

            _showHandler.ElementId = fi.Id;
            var res = _showEvent.Raise();            

            if (res == ExternalEventRequest.Denied)
                TaskDialog.Show("Ошибка", $"При выполнении опперации произошла ошибка");
#endif
        }

        // ClashPoint. Кнопка «Удалить»
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
#if (Revit2023 || Debug2023)
            if (!(sender is Button b && b.DataContext is FamilyInstance fi)) return;

            _delHandler.ElementId = fi.Id;
            var result = _delEvent.Raise();             

            if (result == ExternalEventRequest.Accepted)  
                MonolithClashPoints.Remove(fi);       
            else
                TaskDialog.Show("Ошибка", $"При попытке удаления элемента произошла ошибка");
#endif
        }

        // Необработанный элемент. Кнопка «Выбрать элемент»
        private void SelectSkippedElement_Click(object sender, RoutedEventArgs e)
        {
#if (Revit2023 || Debug2023)
            if (sender is Button btn && btn.Tag is Element el)
            {
                _uiDoc.Selection.SetElementIds(new[] { el.Id });
            }
#endif
        }

    }

}
