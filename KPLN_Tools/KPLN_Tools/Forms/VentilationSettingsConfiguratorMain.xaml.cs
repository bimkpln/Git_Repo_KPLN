using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Data;
using System.Windows.Input;
using Command = KPLN_Tools.ExternalCommands.Command_VentilationSettingsConfigurator;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using Document = Autodesk.Revit.DB.Document;

namespace KPLN_Tools.Forms
{
    public partial class VentilationSettingsConfiguratorMain : Window
    {
        private readonly Command.FamilyRequestHandler _handler;
        private ExternalEvent _externalEvent;
        private bool _isBusy;
        private string _lastSavedPath;
        private string _workingFamilyPath;
        private string _workingSourcePath;
        private string _familyContextKey;
        private sealed class WorkspaceState
        {
            internal Command.FamilyPackage Family;
            internal Command.FamilyTypeItem Selected;
        }
        private readonly Dictionary<string, WorkspaceState> _workspaces = new Dictionary<string, WorkspaceState>(StringComparer.OrdinalIgnoreCase);
        private string _errorDetails;
        private bool _hasOperationError;
        private readonly ObservableCollection<Command.FamilyTypeItem> _types = new ObservableCollection<Command.FamilyTypeItem>();
        private Command.FamilyTypeItem _currentType;
        private Command.SectionCatalog _sectionCatalog;
        private bool _switchingType;
        private Command.InstallationConfiguration _configuration;
        private Command.SectionValue _selectedSection;
        private readonly List<Command.SectionValue> _observedSections = new List<Command.SectionValue>();
        private readonly List<Command.DimensionValue> _observedDimensions = new List<Command.DimensionValue>();
        private Command.InstallationInfo _observedInfo;
        private bool _updatingCalculation;
        private TextBlock _installationLengthText;
        private WrapPanel _sectionParametersPanel;
        private readonly List<Command.BooleanValue> _observedBooleans = new List<Command.BooleanValue>();
        private Point _dragStart;
        internal bool IsBusy { get { return _isBusy; } }
        internal bool HasOperationError { get { return _hasOperationError; } }
        internal void CreateTypeIfReady() { if (_sectionCatalog != null) CreateType(); }

        // Конструктор вызывается только из IExternalCommand.Execute — в контексте Revit API.
        public VentilationSettingsConfiguratorMain(UIApplication uiapp, UIDocument uidoc)
        {
            InitializeComponent();
            _types.Add(new Command.FamilyTypeItem { IsCreate = true });
            FamilyTypesListBox.ItemsSource = _types;
            new WindowInteropHelper(this).Owner = uiapp.MainWindowHandle;
            _handler = new Command.FamilyRequestHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            try { _handler.StartTracking(uiapp); }
            catch
            {
                ReleaseExternalEvent();
                throw;
            }
            SetBusy(false);
        }

        internal void SetProject(Command.ProjectMatch match, string openFilePath)
        {
            // Только снимок данных: свойства Revit Document интерфейс не читает.
            string project = match.Project == null ? Command.UnknownProjectName : match.Project.DisplayName;
            ProjectHeaderTextBlock.Text = project + " - " + openFilePath;
            ProjectHeaderTextBlock.ToolTip = ProjectHeaderTextBlock.Text;
        }

        private void FamilyTypesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_switchingType || _isBusy) return;
            var selected = FamilyTypesListBox.SelectedItem as Command.FamilyTypeItem;
            if (selected == null) return;
            if (selected.IsCreate)
            {
                if (_sectionCatalog == null)
                {
                    _switchingType = true;
                    try { FamilyTypesListBox.SelectedItem = _currentType; }
                    finally { _switchingType = false; }
                    QueueRequest(Command.RequestKind.LoadSectionCatalog);
                }
                else CreateType();
                return;
            }
            SelectType(selected);
        }

        private void CreateType()
        {
            var item = Command.FamilyTypeItem.CreateDraft(_types, _sectionCatalog);
            item.SavedPath = _workingFamilyPath; item.SourcePath = _workingSourcePath;
            _types.Add(item);
            _switchingType = true;
            try { FamilyTypesListBox.SelectedItem = item; }
            finally { _switchingType = false; }
            SelectType(item);
        }

        private void SelectType(Command.FamilyTypeItem selected)
        {
            if (selected.Configuration == null)
            {
                _switchingType = true;
                try { FamilyTypesListBox.SelectedItem = _currentType; }
                finally { _switchingType = false; }
                SetStatus(selected.LoadError ?? "Не удалось прочитать настройки типа.", true);
                return;
            }
            if (_currentType != null)
            {
                _currentType.SelectedSectionIndex = Math.Max(0, _configuration.Blocks.IndexOf(_selectedSection));
                _currentType.SelectedTab = ConfigurationTabs.SelectedIndex;
            }
            _currentType = selected;
            _lastSavedPath = selected.SavedPath;
            _switchingType = true;
            try
            {
                SetConfiguration(selected.Configuration);
                TypeNameTextBox.Text = selected.Name;
                ConfigurationTabs.SelectedIndex = Math.Max(0, selected.SelectedTab);
                NoTypeTextBlock.Visibility = Visibility.Collapsed;
                ConfigurationTabs.Visibility = Visibility.Visible;
                RenderBlocks(); RenderSelectedSection();
            }
            finally { _switchingType = false; }
            selected.MarkChanged();
            SetBusy(_isBusy);
            SetStatus("Выбран тип «" + selected.DisplayName + "».", false);
            RecalculateLocally();
        }

        private void TypeName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_switchingType && _currentType != null && _currentType.Name != TypeNameTextBox.Text)
            { _currentType.Name = TypeNameTextBox.Text; _currentType.MarkChanged(); _currentType.HasAutomaticName = false; }
        }

        private void DeleteType_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            var item = (sender as FrameworkElement)?.DataContext as Command.FamilyTypeItem;
            if (_isBusy || item == null || item.IsCreate || !_types.Contains(item)) return;
            bool persisted = !string.IsNullOrWhiteSpace(item.PersistedName);
            string text = persisted ? "Тип «" + item.PersistedName + "» будет удалён из файла семейства. После удаления файл будет сразу сохранён. Продолжить?"
                : "Удалить новый тип «" + item.DisplayName + "»? Он ещё не сохранён в семействе.";
            if (MessageBox.Show(this, text, "KPLN. Конфигуратор вент. установок. Удаление типа.",
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes) return;
            if (!persisted) { RemoveTypeFromList(item); return; }
            try
            {
                RaiseRequest(new Command.FamilyRequest
                {
                    Kind = Command.RequestKind.DeleteType,
                    TargetType = item,
                    TypeName = item.Name,
                    PersistedName = item.PersistedName,
                    OutputPath = item.SavedPath,
                    FamilyPath = item.SourcePath ?? item.SavedPath
                });
            }
            catch (Exception ex) { SetBusy(false); SetOperationError(Command.OperationError.Create("Удаление типа", ex, null)); }
        }

        private void RemoveTypeFromList(Command.FamilyTypeItem item)
        {
            int index = _types.IndexOf(item);
            bool active = ReferenceEquals(item, _currentType);
            _switchingType = true;
            try
            {
                _types.Remove(item);
                if (active)
                {
                    foreach (var block in _observedSections) block.PropertyChanged -= SectionChanged;
                    _observedSections.Clear();
                    ClearDimensionObservers(); ClearInfoObserver();
                    _currentType = null; _configuration = null; _selectedSection = null; _lastSavedPath = null;
                    var next = _types.Skip(Math.Min(index, _types.Count)).FirstOrDefault(t => !t.IsCreate && t.Configuration != null)
                        ?? _types.LastOrDefault(t => !t.IsCreate && t.Configuration != null);
                    FamilyTypesListBox.SelectedItem = next;
                    if (next != null) SelectType(next);
                    else
                    {
                        TypeNameTextBox.Clear(); SectionBlocksPanel.Children.Clear(); SelectedSectionPanel.Children.Clear();
                        InstallationWidthEditor.Content = null; SectionCountComboBox.ItemsSource = null;
                        AddValveCheckBox.IsChecked = false; ConfigurationTabs.Visibility = Visibility.Collapsed;
                        NoTypeTextBlock.Visibility = Visibility.Visible;
                    }
                }
                else FamilyTypesListBox.SelectedItem = _currentType;
            }
            finally { _switchingType = false; }
            SetBusy(false);
            SetStatus("Тип «" + item.DisplayName + "» удалён из списка.", false);
        }

        internal void SetBusy(bool value)
        {
            _isBusy = value;
            EditorPanel.IsEnabled = !value;
            OpenButton.IsEnabled = !value && _configuration != null;
            AddButton.IsEnabled = OpenButton.IsEnabled;
            SaveButton.IsEnabled = OpenButton.IsEnabled;
            CloseButton.IsEnabled = !value;
            ErrorDetailsButton.IsEnabled = !value;
            SectionCountComboBox.IsEnabled = !value && _configuration != null;
            TypeNameTextBox.IsEnabled = !value && _currentType != null;
        }

        internal void SetStatus(string text, bool error)
        {
            string message = string.IsNullOrWhiteSpace(text) ? (error ? "Ошибка операции." : string.Empty) : text;
            StatusTextBlock.Text = error ? "Ошибка" : string.Join(" ", message.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
            StatusTextBlock.ToolTip = message;
            StatusTextBlock.Foreground = error ? Brushes.Firebrick : new SolidColorBrush(Color.FromRgb(48, 71, 94));
            _errorDetails = error ? message : null;
            _hasOperationError = error;
            ErrorDetailsButton.Visibility = error ? Visibility.Visible : Visibility.Collapsed;
        }

        internal void SetOperationError(Command.OperationError error)
        {
            SetStatus(error.Summary, true);
            _errorDetails = error.Details;
            StatusTextBlock.ToolTip = error.Summary;
        }

        private void ErrorDetails_Click(object sender, RoutedEventArgs e)
        {
            if (!_hasOperationError || string.IsNullOrWhiteSpace(_errorDetails)) return;
            string details = _errorDetails;
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
                FontFamily = FontFamily,
                FontSize = FontSize
            };
            var layout = new DockPanel { Margin = new Thickness(16) };
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var copy = new Button { Content = "Копировать", Padding = new Thickness(14, 6, 14, 6) };
            var close = new Button
            {
                Content = "Закрыть",
                IsCancel = true,
                Padding = new Thickness(14, 6, 14, 6),
                Margin = new Thickness(10, 0, 0, 0)
            };
            buttons.Children.Add(copy); buttons.Children.Add(close);
            DockPanel.SetDock(buttons, Dock.Bottom); layout.Children.Add(buttons);
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
                try { Clipboard.SetText(details); copy.Content = "Скопировано"; }
                catch (System.Runtime.InteropServices.ExternalException)
                { report.Focus(); report.SelectAll(); copy.Content = "Текст выделен"; }
            };
            close.Click += (s, args) => dialog.Close();
            layout.Children.Add(report); dialog.Content = layout; dialog.ShowDialog();
        }

        internal string ChooseOutputPath(string projectCode, string familyPath, bool knownProject)
        {
            var dialog = new SaveFileDialog
            {
                Title = knownProject ? "Папка проекта пока не задана — куда сохранить семейство" : "Куда сохранить семейство",
                Filter = "Семейство Revit (*.rfa)|*.rfa",
                DefaultExt = ".rfa",
                AddExtension = true,
                CheckPathExists = true,
                OverwritePrompt = true,
                // Помним папку, но имя заново формируем для текущего проекта / цели загрузки.
                FileName = string.IsNullOrWhiteSpace(familyPath) ? MakeFileName(projectCode) : Path.GetFileName(familyPath)
            };
            string previous = familyPath ?? _lastSavedPath;
            if (previous != null && Directory.Exists(Path.GetDirectoryName(previous)))
                dialog.InitialDirectory = Path.GetDirectoryName(previous);
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

        internal void RememberSavedFamily(Command.FamilyPackage family, string typeName)
        {
            string previous = _currentType.SavedPath;
            _lastSavedPath = _workingFamilyPath = family.Path; _workingSourcePath = family.SourcePath ?? family.Path;
            _sectionCatalog = family.Catalog;
            _currentType.Name = typeName; _currentType.PersistedName = typeName; _currentType.SavedPath = family.Path; _currentType.SourcePath = family.SourcePath ?? family.Path; _currentType.MarkSaved();
            foreach (var item in _types.Where(t => !t.IsCreate && t != _currentType))
                if (item.SavedPath == null || item.SavedPath == previous && family.Types.Any(t => t.Name == item.PersistedName)) { item.SavedPath = family.Path; item.SourcePath = family.SourcePath ?? family.Path; }
            foreach (var saved in family.Types)
            {
                var existing = _types.FirstOrDefault(t => !t.IsCreate && t.SavedPath == family.Path && t.PersistedName == saved.PersistedName);
                if (existing == null)
                {
                    if (!_types.Any(t => !t.IsCreate && t.Name == saved.Name)) _types.Add(saved);
                }
                else if (!existing.IsDirty && saved.Configuration != null) { existing.Configuration = saved.Configuration; existing.MarkSaved(); }
            }
            SelectType(_currentType);
        }

        internal void RememberOpenTypeIdentity(string path, string typeName)
        {
            if (_currentType == null) return;
            _currentType.Name = typeName; _currentType.PersistedName = typeName;
            _currentType.SavedPath = _currentType.SourcePath = path; _currentType.RequireSave();
        }

        internal void CompleteDeletedFamily(Command.FamilyPackage family, Command.FamilyTypeItem deleted)
        {
            if (deleted != null && _types.Contains(deleted)) RemoveTypeFromList(deleted);
            _workingFamilyPath = _lastSavedPath = family.Path; _workingSourcePath = family.SourcePath ?? family.Path;
            _sectionCatalog = family.Catalog;
            foreach (var item in _types.Where(t => !t.IsCreate))
            {
                item.SavedPath = family.Path; item.SourcePath = family.SourcePath ?? family.Path;
                var saved = family.Types.FirstOrDefault(t => t.PersistedName == item.PersistedName);
                if (!item.IsDirty && saved?.Configuration != null) { item.Configuration = saved.Configuration; item.MarkSaved(); }
            }
            foreach (var saved in family.Types)
                if (!_types.Any(t => !t.IsCreate && t.PersistedName == saved.PersistedName)) _types.Add(saved);
            if (_currentType != null && _currentType.Configuration != null) SelectType(_currentType);
        }

        internal bool SwitchFamilyContext(string key)
        {
            if (_familyContextKey == key) return true;
            if (_familyContextKey != null)
                _workspaces[_familyContextKey] = new WorkspaceState
                {
                    Family = new Command.FamilyPackage
                    {
                        Path = _workingFamilyPath,
                        SourcePath = _workingSourcePath,
                        Catalog = _sectionCatalog,
                        Types = _types.Where(t => !t.IsCreate).ToList()
                    },
                    Selected = _currentType
                };
            _familyContextKey = key;
            WorkspaceState workspace;
            if (_workspaces.TryGetValue(key, out workspace))
            {
                SetFamily(workspace.Family);
                if (workspace.Selected != null && workspace.Selected.Configuration != null)
                {
                    _switchingType = true;
                    try { FamilyTypesListBox.SelectedItem = workspace.Selected; }
                    finally { _switchingType = false; }
                    SelectType(workspace.Selected);
                }
                return true;
            }
            SetFamily(new Command.FamilyPackage());
            return false;
        }

        internal string AvailableAutomaticName(string requested, IEnumerable<string> fileNames)
        {
            var names = fileNames.Concat(_types.Where(t => !t.IsCreate && t != _currentType).Select(t => t.Name))
                .ToList();
            return Command.FamilyTypeItem.AvailableAutomaticName(requested, names);
        }

        internal void MergeDiscoveredFamily(Command.FamilyPackage family, string currentName)
        {
            if (_currentType != null && _currentType.Name != currentName)
            {
                _switchingType = true;
                try { _currentType.Name = currentName; TypeNameTextBox.Text = currentName; }
                finally { _switchingType = false; }
            }
            foreach (var saved in family.Types)
            {
                var existing = _types.FirstOrDefault(t => !t.IsCreate && t.PersistedName == saved.PersistedName
                    && (string.IsNullOrWhiteSpace(t.SavedPath) || string.Equals(t.SavedPath, saved.SavedPath, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(t.SourcePath, _workingSourcePath, StringComparison.OrdinalIgnoreCase)));
                if (existing == null) _types.Add(saved);
                else
                {
                    existing.SavedPath = saved.SavedPath; existing.SourcePath = saved.SourcePath;
                    if (existing != _currentType && !existing.IsDirty && saved.Configuration != null)
                    { existing.Configuration = saved.Configuration; existing.MarkSaved(); }
                }
            }
        }

        internal void SetFamily(Command.FamilyPackage family)
        {
            _switchingType = true;
            try
            {
                foreach (var block in _observedSections) block.PropertyChanged -= SectionChanged;
                _observedSections.Clear(); ClearDimensionObservers(); ClearInfoObserver();
                _types.Clear(); _types.Add(new Command.FamilyTypeItem { IsCreate = true });
                _currentType = null; _configuration = null; _selectedSection = null;
                _sectionCatalog = family.Catalog; _lastSavedPath = _workingFamilyPath = family.Path; _workingSourcePath = family.SourcePath ?? family.Path;
                foreach (var item in family.Types) _types.Add(item);
                TypeNameTextBox.Clear(); SectionBlocksPanel.Children.Clear(); SelectedSectionPanel.Children.Clear();
                ConfigurationTabs.Visibility = Visibility.Collapsed; NoTypeTextBlock.Visibility = Visibility.Visible;
            }
            finally { _switchingType = false; }
            var first = family.Types.FirstOrDefault(t => t.Configuration != null && t.PersistedName == family.SelectedTypeName)
                ?? family.Types.FirstOrDefault(t => t.Configuration != null);
            if (first == null)
            { SetBusy(_isBusy); SetStatus("Выберите «Создать тип», чтобы настроить новую установку или выберите уже имеющийся тип.", false); }
            else
            {
                _switchingType = true;
                try { FamilyTypesListBox.SelectedItem = first; }
                finally { _switchingType = false; }
                SelectType(first);
            }
        }

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
        private void Save_Click(object sender, RoutedEventArgs e) { QueueRequest(Command.RequestKind.Save); }
        private void AddToProject_Click(object sender, RoutedEventArgs e) { QueueRequest(Command.RequestKind.AddToProject); }
        private void Close_Click(object sender, RoutedEventArgs e) { Close(); }

        private void SetConfiguration(Command.InstallationConfiguration configuration)
        {
            _configuration = configuration;
            InfoPanel.DataContext = configuration.Info;
            ClearInfoObserver();
            _observedInfo = configuration.Info;
            _observedInfo.PropertyChanged += InfoChanged;
            var overall = new UniformGrid { Columns = 3 };
            var width = DimensionEditor("Ширина установки", configuration.InstallationWidth, "Установка_Ширина");
            var height = DimensionEditor("Высота установки", configuration.InstallationHeight, "Установка_Высота");
            width.Margin = height.Margin = new Thickness(0, 0, 16, 0);
            overall.Children.Add(width); overall.Children.Add(height);
            var length = new StackPanel();
            length.Children.Add(new TextBlock { Text = "Длина установки", FontSize = 12, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
            _installationLengthText = new TextBlock
            {
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Height = 34,
                Padding = new Thickness(0, 8, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            length.Children.Add(_installationLengthText); overall.Children.Add(length);
            InstallationWidthEditor.Content = overall;
            var additional = new UniformGrid { Columns = 4 };
            additional.Children.Add(BooleanEditor("Рама", configuration.SharedBooleans[Command.SharedParameters.FrameVisible], Command.SharedParameters.FrameVisible, false));
            additional.Children.Add(DimensionEditor("Высота рамы", configuration.FrameHeight, "Основание_Рама_Высота"));
            additional.Children.Add(BooleanEditor("Зона обслуживания", configuration.SharedBooleans[Command.SharedParameters.ServiceRight], Command.SharedParameters.ServiceRight, true));
            additional.Children.Add(DimensionEditor("Зона обслуживания: глубина", configuration.SharedDimensions[Command.SharedParameters.ServiceDepth], Command.SharedParameters.ServiceDepth));
            foreach (FrameworkElement field in additional.Children) field.Margin = new Thickness(0, 0, 16, 0);
            AdditionalInstallationEditor.Content = additional;
            int index = Math.Max(0, Math.Min(_currentType.SelectedSectionIndex, configuration.Blocks.Count - 1));
            _selectedSection = configuration.Blocks[index];
            SyncSectionControls();
            ObserveSections();
        }

        private static Binding ValueBinding(object source, string property)
        { return new Binding(property) { Source = source, Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }; }

        private void ObserveSections()
        {
            foreach (var block in _observedSections) block.PropertyChanged -= SectionChanged;
            _observedSections.Clear();
            foreach (var block in _configuration.Blocks)
            { block.PropertyChanged += SectionChanged; _observedSections.Add(block); }
            ClearDimensionObservers();
            foreach (var dimension in _configuration.Dimensions().Values)
            { dimension.PropertyChanged += DimensionChanged; _observedDimensions.Add(dimension); }
            foreach (var flag in _configuration.SharedBooleans.Values)
            { flag.PropertyChanged += BooleanChanged; _observedBooleans.Add(flag); }
        }

        private void ClearDimensionObservers()
        {
            foreach (var dimension in _observedDimensions) dimension.PropertyChanged -= DimensionChanged;
            _observedDimensions.Clear();
            foreach (var flag in _observedBooleans) flag.PropertyChanged -= BooleanChanged;
            _observedBooleans.Clear();
        }

        private void ClearInfoObserver()
        {
            if (_observedInfo != null) _observedInfo.PropertyChanged -= InfoChanged;
            _observedInfo = null;
        }

        private void InfoChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_switchingType || _currentType == null) return;
            if (e.PropertyName == nameof(Command.InstallationInfo.SystemName) || e.PropertyName == nameof(Command.InstallationInfo.Mark))
            {
                _currentType.UpdateAutomaticName(_types.Where(t => !t.IsCreate && t != _currentType).Select(t => t.Name));
                _switchingType = true;
                try { TypeNameTextBox.Text = _currentType.Name; }
                finally { _switchingType = false; }
            }
            _currentType.MarkChanged();
        }

        private void SectionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!_switchingType && _currentType != null) _currentType.MarkChanged();
            RecalculateLocally();
            if (!_switchingType && ReferenceEquals(sender, _selectedSection) && e.PropertyName == "Type") RenderSectionParameters();
        }

        private void BooleanChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "Value" || _updatingCalculation || _switchingType || _currentType == null) return;
            _currentType.MarkChanged(); RecalculateLocally();
        }

        private void DimensionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Text" && !_updatingCalculation && !((Command.DimensionValue)sender).IsCalculated)
            { if (!_switchingType && _currentType != null) _currentType.MarkChanged(); RecalculateLocally(); }
        }

        private void RecalculateLocally()
        {
            if (_configuration == null || _switchingType || _updatingCalculation) return;
            _updatingCalculation = true;
            try { _configuration.RecalculateLocal(); RenderBlocks(); }
            finally { _updatingCalculation = false; }
        }

        internal void ApplyCalculatedDimensions(Command.InstallationConfiguration calculated)
        {
            if (_configuration == null) return;
            var targets = _configuration.Dimensions();
            _updatingCalculation = true;
            try
            {
                foreach (var pair in calculated.Dimensions())
                {
                    Command.DimensionValue target;
                    if (!targets.TryGetValue(pair.Key, out target)) continue;
                    target.SetCalculatedMode(pair.Value.IsCalculated);
                    if (pair.Value.IsCalculated && pair.Value.IsCurrent) target.ReadResult(pair.Value.Millimeters(pair.Key));
                }
                foreach (var pair in calculated.SharedBooleans)
                {
                    var target = _configuration.SharedBooleans[pair.Key];
                    target.SetCalculatedMode(pair.Value.IsCalculated);
                    if (pair.Value.IsCalculated) target.ReadResult(pair.Value.Value);
                }
            }
            finally { _updatingCalculation = false; }
        }

        private void SectionCount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_switchingType || _isBusy || _configuration == null || SectionCountComboBox.SelectedItem == null) return;
            int count = (int)SectionCountComboBox.SelectedItem;
            if (count != _configuration.SectionCount) EditSections(() => _configuration.Resize(count));
        }

        private void SyncSectionControls()
        {
            bool previous = _switchingType;
            _switchingType = true;
            try
            {
                SectionCountComboBox.ItemsSource = Enumerable.Range(_configuration.Minimum, _configuration.Capacity - _configuration.Minimum + 1).ToList();
                SectionCountComboBox.SelectedItem = _configuration.SectionCount;
                SectionCountComboBox.ToolTip = ParameterTooltip("Секции_Промежуточные_Количество", nameOnly: true);
                AddValveCheckBox.IsChecked = _configuration.HasValve;
                AddValveCheckBox.IsEnabled = _configuration.CanChangeValve;
                AddValveCheckBox.ToolTip = null;
            }
            finally { _switchingType = previous; }
        }

        private void AddValve_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (_switchingType || _isBusy || _configuration == null) return;
            bool enabled = AddValveCheckBox.IsChecked == true;
            EditSections(() =>
            {
                _configuration.SetValve(enabled);
                if (enabled) _selectedSection = _configuration.Blocks[1];
            });
        }

        private void EditSections(Action action)
        {
            try
            {
                action();
                _currentType.MarkChanged();
                SyncSectionControls();
                if (_selectedSection == null || !_configuration.Blocks.Contains(_selectedSection))
                    _selectedSection = _configuration.Blocks[_configuration.FirstSectionIndex];
                ObserveSections(); RenderBlocks(); RenderSelectedSection(); SetBusy(false);
                SetStatus("Количество секций: " + _configuration.SectionCount + ".", false);
                RecalculateLocally();
            }
            catch (Exception ex) { SyncSectionControls(); SetStatus(ex.Message, true); }
        }

        private void RenderBlocks()
        {
            if (_configuration == null) return;
            CompositionSummaryTextBlock.Text = "Состав установки";
            CompositionSummaryTextBlock.ToolTip = null;
            if (_installationLengthText != null)
            {
                _installationLengthText.Text = Command.LocalGeometry.Format(_configuration.Geometry.TotalLengthMm) + " мм";
                _installationLengthText.ToolTip = null;
            }
            SectionBlocksPanel.Children.Clear();
            for (int i = 0; i < _configuration.Blocks.Count; i++)
            {
                int index = i;
                int slot = _configuration.SlotAt(i);
                var block = _configuration.Blocks[i];
                var content = new Grid();
                content.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                content.RowDefinitions.Add(new RowDefinition { Height = new GridLength(38) });
                string title = _configuration.BlockTitle(i);
                var label = new StackPanel();
                label.Children.Add(new TextBlock { Text = title, FontSize = 10, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
                label.Children.Add(new TextBlock
                {
                    Text = block.Type == null ? "Состав не задан" : block.Type.TypeName,
                    TextWrapping = TextWrapping.Wrap,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Height = 35
                });
                var select = new Button
                {
                    Content = label,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    HorizontalContentAlignment = HorizontalAlignment.Stretch,
                    Padding = new Thickness(3),
                    MinHeight = 0
                };
                select.Click += (s, e) => { _selectedSection = block; RenderBlocks(); RenderSelectedSection(); };
                if (!block.IsFixed)
                {
                    select.PreviewMouseLeftButtonDown += (s, e) => _dragStart = e.GetPosition(this);
                    select.PreviewMouseMove += (s, e) =>
                    {
                        Point point = e.GetPosition(this);
                        if (e.LeftButton == MouseButtonState.Pressed &&
                            (Math.Abs(point.X - _dragStart.X) > SystemParameters.MinimumHorizontalDragDistance ||
                             Math.Abs(point.Y - _dragStart.Y) > SystemParameters.MinimumVerticalDragDistance))
                        { DragDrop.DoDragDrop(select, new DataObject("KPLN.Ventilation.Section", block), DragDropEffects.Move); e.Handled = true; }
                    };
                }
                content.Children.Add(select);
                var buttons = new UniformGrid { Rows = 2, Columns = 2 };
                buttons.Children.Add(BlockButton("←", "Переместить влево", !block.IsFixed && index > _configuration.FirstSectionIndex,
                    () => EditSections(() => _configuration.Move(index, index - 1))));
                buttons.Children.Add(BlockButton("→", "Переместить вправо", !block.IsFixed && index < _configuration.Blocks.Count - 2,
                    () => EditSections(() => _configuration.Move(index, index + 1))));
                buttons.Children.Add(BlockButton("+", "Добавить секцию", _configuration.SectionCount < _configuration.Capacity,
                    () => EditSections(() => _configuration.InsertAfter(index))));
                buttons.Children.Add(BlockButton("×", "Удалить секцию", !block.IsFixed && _configuration.SectionCount > _configuration.Minimum,
                    () => EditSections(() => _configuration.Remove(index))));
                Grid.SetRow(buttons, 1); content.Children.Add(buttons);
                var border = new Border
                {
                    Child = content,
                    Height = 116,
                    Padding = new Thickness(2),
                    Margin = new Thickness(2, 0, 2, 0),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    BorderBrush = ReferenceEquals(block, _selectedSection) ? new SolidColorBrush(Color.FromRgb(36, 79, 120)) : Brushes.LightSlateGray,
                    Background = block.IsValve ? new SolidColorBrush(Color.FromRgb(237, 247, 247))
                        : block.IsConnector ? new SolidColorBrush(Color.FromRgb(237, 243, 248)) : Brushes.White,
                    AllowDrop = !block.IsFixed
                };
                border.Drop += (s, e) =>
                {
                    var moved = e.Data.GetData("KPLN.Ventilation.Section") as Command.SectionValue;
                    if (moved == null) return;
                    EditSections(() => _configuration.Move(_configuration.Blocks.IndexOf(moved), index));
                    e.Handled = true;
                };
                var dimensions = new Grid { Height = 16, Margin = new Thickness(3, 0, 3, 4) };
                var length = DimensionCaption("Длина", block.Length);
                dimensions.Children.Add(length);
                var tooltip = new StackPanel();
                tooltip.Children.Add(new TextBlock { Text = _configuration.Catalog[slot].DisplayName });
                tooltip.Children.Add(new TextBlock { Text = block.Type?.DisplayName, Margin = new Thickness(0, 2, 0, 6) });
                tooltip.Children.Add(DimensionCaption("Ширина", _configuration.BlockWidth(i)));
                tooltip.Children.Add(DimensionCaption("Высота", _configuration.BlockHeight(i)));
                tooltip.Children.Add(DimensionCaption("Длина", block.Length));
                if (block.IsConnector)
                {
                    tooltip.Children.Add(DimensionCaption("Смещение по X", block.OffsetX));
                    tooltip.Children.Add(DimensionCaption("Смещение по Y", block.OffsetY));
                }
                tooltip.Children.Add(new TextBlock { Text = "Размеры в мм", Foreground = Brushes.SlateGray, FontSize = 10 });
                border.ToolTip = tooltip;
                var cell = new StackPanel(); cell.Children.Add(dimensions); cell.Children.Add(border);
                SectionBlocksPanel.Children.Add(cell);
            }
        }

        private static TextBlock DimensionCaption(string label, Command.DimensionValue value)
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.SlateGray));
            var invalid = new DataTrigger { Binding = new Binding("IsValid") { Source = value }, Value = false };
            invalid.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.Firebrick)); style.Triggers.Add(invalid);
            var text = new TextBlock { FontSize = 10, Style = style, TextTrimming = TextTrimming.CharacterEllipsis };
            text.SetBinding(TextBlock.TextProperty, new Binding("Display") { Source = value, StringFormat = label + " - {0}" });
            text.SetBinding(TextBlock.ToolTipProperty, new Binding("Display") { Source = value, StringFormat = label + " - {0} мм" });
            return text;
        }

        private FrameworkElement DimensionEditor(string label, Command.DimensionValue value, string parameterName)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 12,
                Foreground = Brushes.SlateGray,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 5)
            });
            var tooltip = ParameterTooltip(parameterName, nameOnly: true);
            var input = new TextBox
            {
                DataContext = value,
                Style = (Style)FindResource("DimensionTextBoxStyle"),
                ToolTip = tooltip
            };
            input.SetBinding(TextBox.TextProperty, ValueBinding(value, "Text"));
            input.SetBinding(TextBox.IsReadOnlyProperty, new Binding("IsReadOnly") { Source = value });
            panel.Children.Add(input);
            return panel;
        }

        private FrameworkElement BooleanEditor(string label, Command.BooleanValue value, string parameterName, bool side)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
            FrameworkElement input;
            if (side)
            {
                var combo = new ComboBox { ItemsSource = new[] { "Справа", "Слева" }, MinHeight = 34, Style = (Style)FindResource("CleanComboBoxStyle") };
                combo.SetBinding(ComboBox.SelectedIndexProperty, ValueBinding(value, "SideIndex"));
                input = combo;
            }
            else
            {
                var check = new CheckBox { Content = "Показывать раму", MinHeight = 34, VerticalContentAlignment = VerticalAlignment.Center };
                check.SetBinding(CheckBox.IsCheckedProperty, ValueBinding(value, "Value"));
                input = check;
            }
            input.ToolTip = ParameterTooltip(parameterName, nameOnly: true);
            input.SetBinding(UIElement.IsEnabledProperty, new Binding("IsEnabled") { Source = value });
            panel.Children.Add(input);
            return panel;
        }

        private static FrameworkElement DimensionInfo(string label, Command.DimensionValue value)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
            var text = new TextBlock { FontSize = 14, Height = 34, Padding = new Thickness(0, 8, 0, 0) };
            text.SetBinding(TextBlock.TextProperty, new Binding("Display") { Source = value });
            panel.Children.Add(text);
            return panel;
        }

        private ToolTip ParameterTooltip(string name, bool nameOnly = false)
        {
            var dependencies = _configuration?.Catalog.Dependencies;
            string text = nameOnly ? dependencies?.Find(name)?.Name ?? name
                : dependencies == null ? name : dependencies.Tooltip(name);
            return new ToolTip { Content = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, MaxWidth = 740 } };
        }

        private Button BlockButton(string caption, string tooltip, bool enabled, Action action)
        {
            string geometry = caption == "+" ? "M 5 0 L 5 10 M 0 5 L 10 5" : caption == "×" ? "M 1 1 L 9 9 M 9 1 L 1 9"
                : caption == "←" ? "M 9 5 L 1 5 M 4 2 L 1 5 L 4 8" : "M 1 5 L 9 5 M 6 2 L 9 5 L 6 8";
            var brush = new SolidColorBrush(caption == "+" ? Color.FromRgb(40, 132, 83)
                : caption == "×" ? Color.FromRgb(195, 63, 69) : Color.FromRgb(83, 101, 122));
            var icon = new System.Windows.Shapes.Path
            {
                Data = Geometry.Parse(geometry),
                Stroke = brush,
                StrokeThickness = 1.6,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                Width = 10,
                Height = 10
            };
            var button = new Button
            {
                Content = icon,
                ToolTip = tooltip,
                IsEnabled = enabled,
                Margin = new Thickness(1),
                Style = (Style)FindResource("IconButtonStyle")
            };
            System.Windows.Automation.AutomationProperties.SetName(button, tooltip);
            button.Click += (s, e) => { action(); e.Handled = true; };
            return button;
        }

        private void RenderSelectedSection()
        {
            SelectedSectionPanel.Children.Clear();
            if (_selectedSection == null) return;
            int index = _configuration.Blocks.IndexOf(_selectedSection);
            int slot = _configuration.SlotAt(index);
            SelectedSectionPanel.Children.Add(new TextBlock
            {
                Text = _selectedSection.IsValve ? "Клапан" : _selectedSection.IsConnector ? "Соединитель" : "Секция",
                FontWeight = FontWeights.SemiBold,
                FontSize = 15,
                Margin = new Thickness(0, 0, 0, 10)
            });
            var definition = _configuration.Catalog[slot];
            SelectedSectionPanel.Children.Add(new TextBlock
            {
                Text = definition.DisplayName,
                ToolTip = ParameterTooltip(definition.ParameterName),
                TextWrapping = TextWrapping.Wrap,
                FontSize = 11,
                Foreground = Brushes.SlateGray,
                Margin = new Thickness(0, 0, 0, 8)
            });
            // После перемещения секции используем объект из списка именно текущего параметра.
            var choices = definition.Choices.Where(t => !_selectedSection.IsConnector || t.IsConnectorType).ToList();
            if (_selectedSection.Type != null)
            {
                if (!choices.Any(t => t.Key == _selectedSection.Type.Key)) choices.Insert(0, _selectedSection.Type);
                _selectedSection.Type = choices.Single(t => t.Key == _selectedSection.Type.Key);
            }
            var selector = new ComboBox
            {
                ItemsSource = choices,
                MinHeight = 54,
                Style = (Style)FindResource("CleanComboBoxStyle"),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment = VerticalAlignment.Center,
                ItemTemplate = (DataTemplate)FindResource("SectionTypeTemplate"),
                MaxDropDownHeight = 360,
                IsTextSearchEnabled = true
            };
            TextSearch.SetTextPath(selector, "TypeName");
            ScrollViewer.SetHorizontalScrollBarVisibility(selector, ScrollBarVisibility.Disabled);
            selector.SetBinding(ComboBox.SelectedItemProperty, ValueBinding(_selectedSection, "Type"));
            selector.SetBinding(ComboBox.ToolTipProperty, new Binding("Type.DisplayName") { Source = _selectedSection });
            SelectedSectionPanel.Children.Add(selector);
            var sizes = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 12, 0, 0) };
            var length = DimensionEditor("Длина", _selectedSection.Length, Command.InstallationConfiguration.DimensionParameterName(slot, "Длина"));
            length.Width = 150; sizes.Children.Add(length);
            if (_selectedSection.IsConnector)
            {
                var width = DimensionEditor("Ширина", _selectedSection.Width, Command.InstallationConfiguration.DimensionParameterName(slot, "Ширина"));
                var height = DimensionEditor("Высота", _selectedSection.Height, Command.InstallationConfiguration.DimensionParameterName(slot, "Высота"));
                var offsetX = DimensionEditor("Смещение по X", _selectedSection.OffsetX, Command.InstallationConfiguration.DimensionParameterName(slot, "Смещение по X"));
                var offsetY = DimensionEditor("Смещение по Y", _selectedSection.OffsetY, Command.InstallationConfiguration.DimensionParameterName(slot, "Смещение по Y"));
                foreach (var field in new[] { width, height, offsetX, offsetY }) { field.Width = 150; field.Margin = new Thickness(12, 0, 0, 0); }
                sizes.Children.Add(width); sizes.Children.Add(height);
                sizes.Children.Add(offsetX); sizes.Children.Add(offsetY);
            }
            else
            {
                var width = DimensionInfo("Ширина", _configuration.BlockWidth(index));
                var height = DimensionInfo("Высота", _configuration.BlockHeight(index));
                width.Width = height.Width = 150; width.Margin = height.Margin = new Thickness(12, 0, 0, 0);
                sizes.Children.Add(width); sizes.Children.Add(height);
            }
            SelectedSectionPanel.Children.Add(sizes);
            _sectionParametersPanel = new WrapPanel { Margin = new Thickness(0, 14, 0, 0) };
            SelectedSectionPanel.Children.Add(_sectionParametersPanel);
            RenderSectionParameters();
        }

        private void RenderSectionParameters()
        {
            if (_sectionParametersPanel == null) return;
            _sectionParametersPanel.Children.Clear();
            if (_selectedSection == null) return;
            var names = Command.SharedParameters.ForSection(_selectedSection.Type).ToList();
            _sectionParametersPanel.Visibility = names.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
            foreach (string name in names)
            {
                Command.BooleanValue flag;
                FrameworkElement editor = _configuration.SharedBooleans.TryGetValue(name, out flag)
                    ? BooleanEditor(Command.SharedParameters.Label(name), flag, name, true)
                    : DimensionEditor(Command.SharedParameters.Label(name), _configuration.SharedDimensions[name], name);
                editor.Width = 220; editor.Margin = new Thickness(0, 0, 14, 12);
                _sectionParametersPanel.Children.Add(editor);
            }
        }

        private void QueueRequest(Command.RequestKind kind)
        {
            if (_isBusy || _externalEvent == null) return;
            try
            {
                string typeName = null;
                Command.InstallationConfiguration snapshot = null;
                if (kind != Command.RequestKind.LoadSectionCatalog)
                {
                    if (_configuration == null || _currentType == null) throw new InvalidOperationException("Сначала создайте тип.");
                    var missing = _configuration.IncompleteFields();
                    if (string.IsNullOrWhiteSpace(TypeNameTextBox.Text)) missing.Insert(0, "Имя типа");
                    if (missing.Count > 0)
                    {
                        MessageBox.Show(this, "Проверьте состав и параметры установки:\n\n"
                            + string.Join("\n", missing.Select(name => "• " + name))
                            + "\n\nРазмеры должны быть больше нуля; смещения могут быть нулевыми или отрицательными.",
                            "KPLN. Конфигуратор вент. установок. Не заполнены параметры.", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    _currentType.MarkChanged();
                    if (kind != Command.RequestKind.Save && (_currentType.IsDirty || string.IsNullOrWhiteSpace(_currentType.SavedPath)
                        || string.IsNullOrWhiteSpace(_currentType.PersistedName)))
                    {
                        MessageBox.Show(this, "Сначала сохраните тип и изменения кнопкой «Сохранить».",
                            "KPLN. Конфигуратор вент. установок", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    snapshot = _configuration.Copy(); snapshot.ValidateForOutput();
                    typeName = Command.FamilyRequestHandler.ValidateTypeName(TypeNameTextBox.Text);
                    if (_types.Any(t => !t.IsCreate && t != _currentType && (_currentType.PersistedName == null || t.PersistedName != null) && string.Equals((t.Name ?? "").Trim(), typeName, StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException("Тип с таким именем уже есть в списке. Укажите другое имя.");
                }
                RaiseRequest(new Command.FamilyRequest
                {
                    Kind = kind,
                    TypeName = typeName,
                    Configuration = snapshot,
                    FamilyPath = _currentType?.SourcePath ?? _currentType?.SavedPath,
                    OutputPath = _currentType?.SavedPath,
                    PersistedName = _currentType?.PersistedName,
                    HasAutomaticName = _currentType != null && _currentType.HasAutomaticName,
                    IsDirty = _currentType != null && _currentType.IsDirty
                });
            }
            catch (Exception ex)
            {
                _handler.PendingRequest = null;
                SetBusy(false);
                SetOperationError(Command.OperationError.Create("Передача запроса Revit", ex, null));
            }
        }

        private void RaiseRequest(Command.FamilyRequest request)
        {
            if (_externalEvent == null) throw new InvalidOperationException("Окно конфигуратора уже закрыто.");
            _handler.PendingRequest = request;
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

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!_isBusy) return;
            e.Cancel = true;
            SetStatus("Дождитесь завершения операции. Если Revit занят другой командой, завершите её.", false);
        }

        private void Window_Closed(object sender, EventArgs e) { ReleaseExternalEvent(); }

        internal void ReleaseExternalEvent()
        {
            ClearDimensionObservers(); ClearInfoObserver();
            // Закрытие WPF-окна происходит вне API-контекста. Отписку выполнит сам Idling.
            _handler.RequestStopTracking();
            if (_externalEvent == null) return;
            _externalEvent.Dispose();
            _externalEvent = null;
        }
    }
}