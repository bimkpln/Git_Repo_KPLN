using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Parameters_Ribbon.ExternalEventHandler;
using KPLN_Parameters_Ribbon.Forms.Common;
using KPLN_Parameters_Ribbon.Forms.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class SumParameters : Window
    {
        private static SumParameters _currentInstance;

#if !Debug2020 && !Revit2020
        private readonly ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;
#endif

        public SumParameters(UIApplication uiapp)
        {
            CurrentSumParametersVM = new SumParametersVM(uiapp);

            InitializeComponent();

            DataContext = CurrentSumParametersVM;

#if !Debug2020 && !Revit2020
            BtnUpdate.Visibility = Visibility.Collapsed;

            _selExtEv = FormEventSubscriptionHelper.CreateSelectionChangedEvent(handler =>
            {
                _selHandler = handler;
                _selHandler.CurrentSumParametersVM = CurrentSumParametersVM;
            });

            ExternalEvent unsubSelExtEv = FormEventSubscriptionHelper.CreateSelectionUnsubscribeEvent(OnSelectionChanged);
            FormEventSubscriptionHelper.SubscribeSelectionChanged(uiapp, this, OnSelectionChanged, unsubSelExtEv);
#endif

            _currentInstance = this;
            Closed += (sender, args) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        public SumParametersVM CurrentSumParametersVM { get; set; }

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => _selExtEv?.Raise();
#endif

        private void OnCellPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGridCell cell)
            {
                cell.Focus();
                cell.IsSelected = true;
                cell.ContextMenu = CreateCellContextMenu(cell);
            }
        }

        private ContextMenu CreateCellContextMenu(DataGridCell cell)
        {
            ContextMenu contextMenu = new ContextMenu { PlacementTarget = cell };
            MenuItem copyCellValueItem = new MenuItem { Header = "Копировать значение ячейки" };

            copyCellValueItem.Click += OnCopyCellValueClick;
            contextMenu.Items.Add(copyCellValueItem);

            return contextMenu;
        }

        private static void OnCopyCellValueClick(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem
                && menuItem.Parent is ContextMenu contextMenu
                && contextMenu.PlacementTarget is DataGridCell cell)
            {
                System.Windows.Clipboard.SetText(GetCellText(cell));
            }
        }

        private static string GetCellText(DataGridCell cell)
        {
            if (cell.Column?.GetCellContent(cell.DataContext) is TextBlock textBlock)
                return textBlock.Text ?? string.Empty;

            if (cell.Content is TextBlock contentTextBlock)
                return contentTextBlock.Text ?? string.Empty;

            if (cell.Column?.GetCellContent(cell.DataContext) is ContentControl contentControl)
                return contentControl.Content?.ToString() ?? string.Empty;

            return string.Empty;
        }

        public static bool TryActivateExisting()
        {
            if (_currentInstance == null || !_currentInstance.IsLoaded)
            {
                _currentInstance = null;
                return false;
            }

            if (_currentInstance.WindowState == WindowState.Minimized)
                _currentInstance.WindowState = WindowState.Normal;

            if (!_currentInstance.IsVisible)
                _currentInstance.Show();

            _currentInstance.Activate();
            _currentInstance.Focus();

            return true;
        }
    }
}
