using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Forms.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Library_Schedules.Forms
{
    // Поиск внутри выпадающего списка не изменяет ItemsSource и выбранный параметр.
    public sealed class ScheduleParameterComboBox : ComboBox
    {
        private TextBox _search;
        public ScheduleParameterComboBox()
        {
            IsTextSearchEnabled = false;
            VirtualizingStackPanel.SetIsVirtualizing(this, false);
            ItemContainerGenerator.StatusChanged += (s, e) => FilterItems();
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            var popup = GetTemplateChild("PART_Popup") as System.Windows.Controls.Primitives.Popup;
            if (popup == null || popup.Child == null) return;
            var content = popup.Child;
            popup.Child = null;
            var panel = new DockPanel { LastChildFill = true, Background = System.Windows.Media.Brushes.White };
            _search = new TextBox { Margin = new Thickness(4), Padding = new Thickness(4),
                ToolTip = "Поиск по любой части имени параметра. Очистите строку, чтобы показать все параметры." };
            DockPanel.SetDock(_search, Dock.Top);
            panel.Children.Add(_search);
            panel.Children.Add(content);
            popup.Child = panel;
            _search.TextChanged += (s, e) => FilterItems();
            _search.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    IsDropDownOpen = false;
                    Focus();
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Down)
                {
                    for (int i = 0; i < Items.Count; i++)
                    {
                        var item = ItemContainerGenerator.ContainerFromIndex(i) as ComboBoxItem;
                        if (item != null && item.Visibility == Visibility.Visible) { item.Focus(); break; }
                    }
                    e.Handled = true;
                }
            };
        }

        protected override void OnDropDownOpened(EventArgs e)
        {
            base.OnDropDownOpened(e);
            if (_search == null) return;
            _search.Text = "";
            Dispatcher.BeginInvoke(new Action(() => { FilterItems(); _search.Focus(); }),
                System.Windows.Threading.DispatcherPriority.Input);
        }

        private void FilterItems()
        {
            string query = (_search?.Text ?? "").Trim();
            for (int i = 0; i < Items.Count; i++)
            {
                var container = ItemContainerGenerator.ContainerFromIndex(i) as ComboBoxItem;
                if (container == null) continue;
                var parameter = Items[i] as ScheduleParameterChoice;
                bool visible = query.Length == 0 || (parameter?.Name ?? "").IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
                container.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
                container.IsEnabled = visible;
            }
        }
    }

}
