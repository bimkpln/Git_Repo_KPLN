using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using Command = KPLN_Tools.ExternalCommands.Command_VentilationSettingsConfigurator;
using ComboBox = System.Windows.Controls.ComboBox;
using Document = Autodesk.Revit.DB.Document;

namespace KPLN_Tools.Forms
{
    public partial class VentilationSettingsConfiguratorMain : Window
    {
        private readonly Command.FamilyRequestHandler _handler;
        private ExternalEvent _externalEvent;
        private bool _isBusy;
        private string _lastSavedPath;
        private Command.ProjectItem _activeProject;
        private bool _settingProjects;
        internal bool IsBusy { get { return _isBusy; } }

        // Конструктор вызывается только из IExternalCommand.Execute — в контексте Revit API.
        public VentilationSettingsConfiguratorMain(UIApplication uiapp, UIDocument uidoc)
        {
            InitializeComponent();
            new WindowInteropHelper(this).Owner = uiapp.MainWindowHandle;
            _handler = new Command.FamilyRequestHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            try { _handler.StartTracking(uiapp); }
            catch
            {
                ReleaseExternalEvent();
                throw;
            }
        }

        internal void SetProjects(IList<Command.ProjectItem> projects, Command.ProjectMatch match,
            string openFilePath)
        {
            // Только снимки данных: свойств Revit Document интерфейс не читает.
            var items = projects.ToList();
            _activeProject = match.Project;
            if (_activeProject == null)
            {
                _activeProject = new Command.ProjectItem { IsUnknown = true };
                items.Insert(0, _activeProject);
            }
            _settingProjects = true;
            try
            {
                ProjectsListBox.ItemsSource = items;
                ProjectsListBox.SelectedItem = _activeProject;
                ProjectsListBox.ScrollIntoView(_activeProject);
            }
            finally { _settingProjects = false; }
            ProjectNameTextBlock.Text = _activeProject.DisplayName;
            ProjectNameTextBlock.ToolTip = _activeProject.DisplayName;
            // Путь открытого файла не зависит от того, нашлась ли запись проекта в базе.
            ProjectPathTextBox.Text = openFilePath;
            ProjectPathTextBox.ToolTip = ProjectPathTextBox.Text;
        }

        private void ProjectsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Пока нет ручной привязки: выделение всегда соответствует активной модели Revit.
            if (!_settingProjects && _activeProject != null &&
                !ReferenceEquals(ProjectsListBox.SelectedItem, _activeProject))
                ProjectsListBox.SelectedItem = _activeProject;
        }

        internal void SetBusy(bool value)
        {
            _isBusy = value;
            EditorPanel.IsEnabled = !value;
            OpenButton.IsEnabled = !value;
            AddButton.IsEnabled = !value;
            CloseButton.IsEnabled = !value;
        }

        internal void SetStatus(string text, bool error)
        {
            StatusTextBlock.Text = text;
            StatusTextBlock.Foreground = error ? Brushes.Firebrick : new SolidColorBrush(Color.FromRgb(48, 71, 94));
        }

        internal string ChooseOutputPath(string projectCode)
        {
            var dialog = new SaveFileDialog
            {
                Title = "Куда сохранить копию семейства",
                Filter = "Семейство Revit (*.rfa)|*.rfa",
                DefaultExt = ".rfa",
                AddExtension = true,
                CheckPathExists = true,
                OverwritePrompt = true,
                // Помним папку, но имя заново формируем для текущего проекта / цели загрузки.
                FileName = MakeFileName(projectCode)
            };
            if (_lastSavedPath != null && Directory.Exists(Path.GetDirectoryName(_lastSavedPath)))
                dialog.InitialDirectory = Path.GetDirectoryName(_lastSavedPath);
            return dialog.ShowDialog(this) == true ? dialog.FileName : null;
        }

        internal static string MakeFileName(string projectCode)
        {
            string code = Command.CleanDatabaseValue(projectCode);
            if (code.Length == 0) code = Command.UnknownProjectName;
            char[] invalid = Path.GetInvalidFileNameChars();
            string safeName = new string(code.Select(c => invalid.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).TrimEnd('.', ' ');
            if (safeName.Length > 80) safeName = safeName.Substring(0, 80).TrimEnd('.', ' ');
            if (safeName.Length == 0) safeName = Command.UnknownProjectName;
            return "550_Универсальная установка_Одноуровневая_(" + safeName + ").rfa";
        }

        internal void RememberSavedPath(string path) { _lastSavedPath = path; }

        internal Command.OpenDocumentItem ChooseLoadProject(IList<Command.OpenDocumentItem> projects, Document preferred)
        {
            // Служебный диалог выбирает открытый документ для загрузки, а не запись базы.
            // Создаётся здесь, чтобы сохранить прежний набор файлов XAML / C#.
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
                FontFamily = FontFamily,
                FontSize = FontSize
            };
            var layout = new Grid { Margin = new Thickness(20) };
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.Children.Add(new TextBlock { Text = "Выберите открытый проект для загрузки типа", Margin = new Thickness(0, 0, 0, 12) });
            var selector = new ComboBox
            {
                ItemsSource = projects,
                DisplayMemberPath = "LoadTargetName",
                SelectedItem = projects.FirstOrDefault(p => ReferenceEquals(p.Document, preferred)) ?? projects.FirstOrDefault(),
                MinHeight = 36,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(8, 5, 8, 5)
            };
            Grid.SetRow(selector, 1);
            layout.Children.Add(selector);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var cancel = new Button { Content = "Отмена", IsCancel = true, MinWidth = 95, Padding = new Thickness(12, 7, 12, 7) };
            var accept = new Button
            {
                Content = "Выбрать",
                IsDefault = true,
                MinWidth = 95,
                Padding = new Thickness(12, 7, 12, 7),
                Margin = new Thickness(10, 0, 0, 0)
            };
            accept.Click += (s, e) => { if (selector.SelectedItem != null) dialog.DialogResult = true; };
            buttons.Children.Add(cancel);
            buttons.Children.Add(accept);
            Grid.SetRow(buttons, 3);
            layout.Children.Add(buttons);
            dialog.Content = layout;
            return dialog.ShowDialog() == true ? selector.SelectedItem as Command.OpenDocumentItem : null;
        }

        private void OpenInRevit_Click(object sender, RoutedEventArgs e) { QueueRequest(Command.RequestKind.OpenInRevit); }
        private void AddToProject_Click(object sender, RoutedEventArgs e) { QueueRequest(Command.RequestKind.AddToProject); }
        private void Close_Click(object sender, RoutedEventArgs e) { Close(); }

        private void QueueRequest(Command.RequestKind kind)
        {
            if (_isBusy || _externalEvent == null) return;
            try
            {
                string typeName = Command.FamilyRequestHandler.ValidateTypeName(TypeNameTextBox.Text);
                _handler.PendingRequest = new Command.FamilyRequest { Kind = kind, TypeName = typeName };
                SetBusy(true);
                SetStatus("Ожидание Revit. Если выполняется другая команда, завершите её.", false);
                ExternalEventRequest result = _externalEvent.Raise();
                if (result != ExternalEventRequest.Accepted)
                {
                    _handler.PendingRequest = null;
                    SetBusy(false);
                    SetStatus("Revit не принял запрос: " + result + ". Завершите текущую команду и повторите действие.", true);
                }
            }
            catch (Exception ex)
            {
                _handler.PendingRequest = null;
                SetBusy(false);
                SetStatus(ex.Message, true);
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!_isBusy) return;
            e.Cancel = true;
            SetStatus("Дождитесь завершения операции. Если Revit занят другой командой, завершите её.", false);
        }

        private void Window_Closed(object sender, EventArgs e) { ReleaseExternalEvent(); }

        internal void ReleaseExternalEvent()
        {
            // Закрытие WPF-окна происходит вне API-контекста. Отписку выполнит сам Idling.
            _handler.RequestStopTracking();
            if (_externalEvent == null) return;
            _externalEvent.Dispose();
            _externalEvent = null;
        }
    }
}
