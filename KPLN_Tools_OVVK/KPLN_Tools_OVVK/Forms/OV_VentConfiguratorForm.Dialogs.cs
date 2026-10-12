using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ComboBox = System.Windows.Controls.ComboBox;
using Document = Autodesk.Revit.DB.Document;
using TextBox = System.Windows.Controls.TextBox;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OV_VentConfiguratorForm : Window
    {
        private void ErrorDetails_Click(object sender, RoutedEventArgs e)
        {
            if (!Model.HasOperationError || string.IsNullOrWhiteSpace(Model.ErrorDetails))
                return;
            string details = Model.ErrorDetails;
            var dialog = new Window
            {
                Title = "KPLN. Конфигуратор вент. установок. Ошибка.",
                Owner = this,
                Width = 860,
                Height = 520,
                MinWidth = 560,
                MinHeight = 320,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ShowInTaskbar = false,
                Background = Background,
                Resources = Resources,
                FontFamily = FontFamily,
                FontSize = FontSize
            };
            var layout = new DockPanel
            {
                Margin = new Thickness(16)
            };
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var copy = new Button
            {
                Content = "Копировать",
                Padding = new Thickness(14, 6, 14, 6)
            };
            var close = new Button
            {
                Content = "Закрыть",
                IsCancel = true,
                Padding = new Thickness(14, 6, 14, 6),
                Margin = new Thickness(10, 0, 0, 0)
            };
            buttons.Children.Add(copy);
            buttons.Children.Add(close);
            DockPanel.SetDock(buttons, Dock.Bottom);
            layout.Children.Add(buttons);
            var report = new TextBox
            {
                Text = details,
                IsReadOnly = true,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(10),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12
            };
            copy.Click += (s, args) =>
            {
                try
                {
                    Clipboard.SetText(details);
                    copy.Content = "Скопировано";
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    report.Focus();
                    report.SelectAll();
                    copy.Content = "Текст выделен";
                }
            };
            close.Click += (s, args) => dialog.Close();
            layout.Children.Add(report);
            dialog.Content = layout;
            dialog.ShowDialog();
        }

        internal string ChooseExistingFamilyPath(string familyName, string failure)
        {
            if (!string.IsNullOrWhiteSpace(failure))
                MessageBox.Show(this, failure + "\n\nВыберите RFA для основы. Типы и заполненные значения будут перенесены из модели. " + "Наличие этих типов в выбранном файле не требуется.", "KPLN. Чтение существующих типов", MessageBoxButton.OK, MessageBoxImage.Information);
            var dialog = new OpenFileDialog
            {
                Title = "Выберите семейство для чтения или переноса типов из модели",
                Filter = "Семейство Revit (*.rfa)|*.rfa",
                DefaultExt = ".rfa",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false
            };
            if (Directory.Exists(ProjectFamiliesRoot))
                dialog.InitialDirectory = ProjectFamiliesRoot;
            return dialog.ShowDialog(this) == true ? dialog.FileName : null;
        }

        internal string ChooseOutputPath(string projectCode, string familyPath)
        {
            var dialog = new SaveFileDialog
            {
                Title = "Куда сохранить семейство",
                Filter = "Семейство Revit (*.rfa)|*.rfa",
                DefaultExt = ".rfa",
                AddExtension = true,
                CheckPathExists = true,
                OverwritePrompt = true,
                // Помним папку, но имя заново формируем для текущего проекта / цели загрузки.
                FileName = string.IsNullOrWhiteSpace(familyPath) ? MakeFileName(projectCode) : Path.GetFileName(familyPath)
            };
            string previous = familyPath ?? Model.LastSavedPath;
            if (previous != null && Directory.Exists(Path.GetDirectoryName(previous)))
                dialog.InitialDirectory = Path.GetDirectoryName(previous);
            return dialog.ShowDialog(this) == true ? dialog.FileName : null;
        }

        internal OpenDocumentItem ChooseLoadProject(IList<OpenDocumentItem> projects, Document preferred)
        {
            // Служебный диалог выбирает открытый документ для загрузки, а не запись базы.
            var dialog = new Window
            {
                Title = "Добавить в проект",
                Owner = this,
                Width = 620,
                Height = 230,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ShowInTaskbar = false,
                Background = Background,
                Resources = Resources,
                FontFamily = FontFamily,
                FontSize = FontSize
            };
            var layout = new Grid
            {
                Margin = new Thickness(20)
            };
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.Children.Add(new TextBlock { Text = "Выберите открытый проект для загрузки типа", Margin = new Thickness(0, 0, 0, 12) });
            var selector = new ComboBox
            {
                Style = (Style)FindResource("CleanComboBoxStyle"),
                ItemsSource = projects,
                DisplayMemberPath = "LoadTargetName",
                SelectedItem = projects.FirstOrDefault(p => ReferenceEquals(p.Document, preferred)) ?? projects.FirstOrDefault(),
                MinHeight = 36,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(8, 5, 8, 5)
            };
            Grid.SetRow(selector, 1);
            layout.Children.Add(selector);
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var cancel = new Button
            {
                Content = "Отмена",
                IsCancel = true,
                MinWidth = 95,
                Padding = new Thickness(12, 7, 12, 7)
            };
            var accept = new Button
            {
                Content = "Выбрать",
                IsDefault = true,
                MinWidth = 95,
                Padding = new Thickness(12, 7, 12, 7),
                Margin = new Thickness(10, 0, 0, 0)
            };
            accept.Click += (s, e) =>
            {
                if (selector.SelectedItem != null)
                    dialog.DialogResult = true;
            };
            buttons.Children.Add(cancel);
            buttons.Children.Add(accept);
            Grid.SetRow(buttons, 3);
            layout.Children.Add(buttons);
            dialog.Content = layout;
            return dialog.ShowDialog() == true ? selector.SelectedItem as OpenDocumentItem : null;
        }
    }
}
