using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_CommandsWheel.Forms
{
    internal class CommandSearchWindow : Window
    {
        private static CommandSearchWindow _current;

        private readonly List<RevitCommandInfo> _commands;
        private readonly Dictionary<string, RevitCommandInfo> _commandsById;
        private readonly UserSettings _settings;
        private readonly RevitCommandExecutor _executor;
        private readonly TextBox _searchBox;
        private readonly StackPanel _contentPanel;
        private RadioButton _unpinnedWheelRadioButton;
        private RadioButton _pinnedWheelRadioButton;
        private CheckBox _wheelCloseButtonCheckBox;
        private TextBox _wheelShortcutTextBox;
        private TextBox _commandSearchShortcutTextBox;
        private TextBlock _wheelShortcutLayoutHintTextBlock;
        private TextBlock _commandSearchShortcutLayoutHintTextBlock;
        private TextBlock _shortcutStatusTextBlock;
        private bool _isUpdatingSettingsControls;

        private enum CommandListKind
        {
            None,
            Wheel,
            Favorites
        }

        internal CommandSearchWindow(IEnumerable<RevitCommandInfo> commands, UserSettings settings, RevitCommandExecutor executor)
        {
            _current = this;

            _commands = commands
                .Where(command => command != null && !string.IsNullOrWhiteSpace(command.Id))
                .OrderBy(command => command.Name)
                .ToList();

            _commandsById = _commands
                .GroupBy(command => command.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            _settings = settings;
            _executor = executor;

            Title = "KPLN. Штурвал команд. Команды";
            Width = 680;
            Height = 740;
            MinWidth = 600;
            MinHeight = 560;
            Background = Brushes.White;
            Focusable = true;

            _searchBox = CreateSearchBox();
            _contentPanel = new StackPanel();
            Content = CreateContent();

            Loaded += delegate
            {
                _searchBox.Focus();
                Keyboard.Focus(_searchBox);
            };

            PreviewKeyDown += delegate (object sender, KeyEventArgs args)
            {
                if (args.Key == Key.Escape)
                {
                    args.Handled = true;
                    Close();
                }
            };
            Closed += delegate
            {
                if (ReferenceEquals(_current, this))
                {
                    _current = null;
                }
            };

            Rebuild();
        }

        internal static bool TryActivateExisting()
        {
            if (_current == null || !_current.IsVisible)
            {
                return false;
            }

            if (_current.WindowState == WindowState.Minimized)
            {
                _current.WindowState = WindowState.Normal;
            }

            _current.Activate();
            return true;
        }

        private UIElement CreateContent()
        {
            TabControl tabControl = new TabControl { Margin = new Thickness(14) };
            tabControl.Items.Add(new TabItem
            {
                Header = "Команды",
                Content = CreateCommandsContent()
            });
            tabControl.Items.Add(new TabItem
            {
                Header = "Настройки",
                Content = CreateSettingsContent()
            });

            return tabControl;
        }

        private UIElement CreateCommandsContent()
        {
            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            root.Children.Add(_searchBox);

            ScrollViewer scrollViewer = new ScrollViewer
            {
                Content = _contentPanel,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            Grid.SetRow(scrollViewer, 1);
            root.Children.Add(scrollViewer);

            return root;
        }

        private UIElement CreateSettingsContent()
        {
            StackPanel panel = new StackPanel
            {
                Margin = new Thickness(18, 18, 18, 16)
            };

            panel.Children.Add(CreateSettingsHeader("Штурвал"));

            _unpinnedWheelRadioButton = new RadioButton
            {
                Content = "Не закреплён",
                Margin = new Thickness(0, 2, 0, 4),
                GroupName = "WheelMode"
            };
            _unpinnedWheelRadioButton.Checked += delegate
            {
                if (_isUpdatingSettingsControls)
                {
                    return;
                }

                _settings.WheelMode = WheelModeNames.Unpinned;
                _settings.IsWheelCloseButtonVisible = false;
                SaveSettingsAndRefresh();
            };
            panel.Children.Add(_unpinnedWheelRadioButton);

            _pinnedWheelRadioButton = new RadioButton
            {
                Content = "Закреплён",
                Margin = new Thickness(0, 0, 0, 10),
                GroupName = "WheelMode"
            };
            _pinnedWheelRadioButton.Checked += delegate
            {
                if (_isUpdatingSettingsControls)
                {
                    return;
                }

                _settings.WheelMode = WheelModeNames.Pinned;
                _settings.IsWheelCloseButtonVisible = true;
                SaveSettingsAndRefresh();
            };
            panel.Children.Add(_pinnedWheelRadioButton);

            _wheelCloseButtonCheckBox = new CheckBox
            {
                Content = "Кнопка закрытия (красный крест)",
                Margin = new Thickness(0, 0, 0, 16)
            };
            _wheelCloseButtonCheckBox.Checked += delegate
            {
                if (_isUpdatingSettingsControls)
                {
                    return;
                }

                _settings.IsWheelCloseButtonVisible = true;
                SaveSettingsAndRefresh();
            };
            _wheelCloseButtonCheckBox.Unchecked += delegate
            {
                if (_isUpdatingSettingsControls)
                {
                    return;
                }

                _settings.IsWheelCloseButtonVisible = false;
                SaveSettingsAndRefresh();
            };
            panel.Children.Add(_wheelCloseButtonCheckBox);

            panel.Children.Add(new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 224, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Margin = new Thickness(0, 0, 0, 16)
            });

            panel.Children.Add(CreateSettingsHeader("Горячие клавиши"));

            string displayedWheelShortcut;
            string displayedSearchShortcut;
            GetDisplayedShortcutValues(
                out displayedWheelShortcut,
                out displayedSearchShortcut);

            Grid shortcutsRow = new Grid
            {
                Margin = new Thickness(0, 2, 0, 10)
            };
            shortcutsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            shortcutsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            shortcutsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            shortcutsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock wheelShortcutLabel = new TextBlock
            {
                Text = "Штурвал",
                Foreground = new SolidColorBrush(Color.FromRgb(70, 70, 70)),
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(wheelShortcutLabel, 0);
            shortcutsRow.Children.Add(wheelShortcutLabel);

            Grid wheelShortcutEditor = CreateShortcutEditor(
                displayedWheelShortcut,
                out _wheelShortcutTextBox,
                out _wheelShortcutLayoutHintTextBlock);
            wheelShortcutEditor.Margin = new Thickness(0, 0, 18, 0);
            Grid.SetColumn(wheelShortcutEditor, 1);
            shortcutsRow.Children.Add(wheelShortcutEditor);

            TextBlock searchShortcutLabel = new TextBlock
            {
                Text = "Команды",
                Foreground = new SolidColorBrush(Color.FromRgb(70, 70, 70)),
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(searchShortcutLabel, 2);
            shortcutsRow.Children.Add(searchShortcutLabel);

            Grid commandSearchShortcutEditor = CreateShortcutEditor(
                displayedSearchShortcut,
                out _commandSearchShortcutTextBox,
                out _commandSearchShortcutLayoutHintTextBlock);
            Grid.SetColumn(commandSearchShortcutEditor, 3);
            shortcutsRow.Children.Add(commandSearchShortcutEditor);
            RefreshShortcutToolTips();

            panel.Children.Add(shortcutsRow);

            panel.Children.Add(new TextBlock
            {
                Text = "Перезаписывает/дополняет XML файл с настройками для всех имеющихся на ПК версий Revit.\nАвтоматически записываются английская раскладка в верхнем регистре и русская в верхнем и нижнем.",
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                Margin = new Thickness(0, 0, 0, 10),
                TextWrapping = TextWrapping.Wrap
            });

            panel.Children.Add(new TextBlock
            {
                Text = "Чтобы изменения вступили в силу, необходимо перезагрузить Revit.",
                Foreground = new SolidColorBrush(Color.FromRgb(95, 95, 95)),
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 10),
                TextWrapping = TextWrapping.Wrap
            });

            Button applyShortcutsButton = new Button
            {
                Content = "Сохранить горячие клавиши",
                HorizontalAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(14, 6, 14, 6),
                Margin = new Thickness(0, 0, 0, 8)
            };
            applyShortcutsButton.Click += delegate { ApplyKeyboardShortcuts(); };
            panel.Children.Add(applyShortcutsButton);

            _shortcutStatusTextBlock = new TextBlock
            {
                Foreground = new SolidColorBrush(Color.FromRgb(55, 105, 65)),
                TextWrapping = TextWrapping.Wrap
            };
            panel.Children.Add(_shortcutStatusTextBlock);

            RefreshSettingsControls();

            return new ScrollViewer
            {
                Content = panel,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
        }

        private TextBlock CreateSettingsHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(70, 70, 70)),
                Margin = new Thickness(0, 0, 0, 8)
            };
        }

        private Grid CreateShortcutEditor(
            string value,
            out TextBox textBox,
            out TextBlock layoutHintTextBlock)
        {
            Grid editor = new Grid();
            textBox = CreateShortcutTextBox(value);
            layoutHintTextBlock = new TextBlock
            {
                Foreground = new SolidColorBrush(Color.FromRgb(155, 155, 155)),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0),
                IsHitTestVisible = false
            };

            editor.Children.Add(textBox);
            editor.Children.Add(layoutHintTextBlock);
            return editor;
        }

        private TextBox CreateShortcutTextBox(string value)
        {
            TextBox textBox = new TextBox
            {
                Text = value ?? string.Empty,
                MinWidth = 120,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Padding = new Thickness(7, 5, 72, 5),
                VerticalContentAlignment = VerticalAlignment.Center
            };

            textBox.TextChanged += delegate { RefreshShortcutToolTips(); };
            ToolTipService.SetInitialShowDelay(textBox, 150);
            ToolTipService.SetShowDuration(textBox, 60000);
            return textBox;
        }

        private void RefreshShortcutToolTips()
        {
            UpdateShortcutToolTip(
                _wheelShortcutTextBox,
                _wheelShortcutLayoutHintTextBlock);
            UpdateShortcutToolTip(
                _commandSearchShortcutTextBox,
                _commandSearchShortcutLayoutHintTextBlock);
        }

        private static void UpdateShortcutToolTip(
            TextBox textBox,
            TextBlock layoutHintTextBlock)
        {
            if (textBox == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.ToolTip = "Горячие клавиши не назначены.";
                if (layoutHintTextBlock != null)
                {
                    layoutHintTextBlock.Text = string.Empty;
                }
                return;
            }

            string preview;
            string englishUpper;
            string russianUpper;
            string error;
            if (!KeyboardShortcutService.TryBuildShortcutPreview(
                textBox.Text,
                out preview,
                out error))
            {
                textBox.ToolTip = "Значение недопустимо.\n" + error;
                if (layoutHintTextBlock != null)
                {
                    layoutHintTextBlock.Text = string.Empty;
                }
                return;
            }

            KeyboardShortcutService.TryNormalizeShortcutInput(
                textBox.Text,
                out englishUpper,
                out russianUpper,
                out error);
            if (layoutHintTextBlock != null)
            {
                layoutHintTextBlock.Text = russianUpper;
            }
            textBox.ToolTip = "Будут назначены: " + preview;
        }

        private void GetDisplayedShortcutValues(
            out string wheelShortcut,
            out string searchShortcut)
        {
            wheelShortcut = _settings.WheelShortcut ?? string.Empty;
            searchShortcut = _settings.CommandSearchShortcut ?? string.Empty;

            string wheelCommandId = KeyboardShortcutService.FindCommandId(
                _commands,
                typeof(ExternalCommands.CommandsWheel));
            string searchCommandId = KeyboardShortcutService.FindCommandId(
                _commands,
                typeof(ExternalCommands.CommandSearch));
            string storedWheelShortcut;
            string storedSearchShortcut;

            if (KeyboardShortcutService.TryReadCurrentShortcuts(
                ModuleData.RevitVersion,
                ModuleData.RevitVersionName,
                wheelCommandId,
                searchCommandId,
                out storedWheelShortcut,
                out storedSearchShortcut))
            {
                if (storedWheelShortcut != null)
                {
                    wheelShortcut = storedWheelShortcut;
                }

                if (storedSearchShortcut != null)
                {
                    searchShortcut = storedSearchShortcut;
                }
            }

            wheelShortcut = NormalizeShortcutForDisplay(wheelShortcut);
            searchShortcut = NormalizeShortcutForDisplay(searchShortcut);
        }

        private static string NormalizeShortcutForDisplay(string value)
        {
            string englishUpper;
            string russianUpper;
            string error;
            return KeyboardShortcutService.TryNormalizeShortcutInput(
                value,
                out englishUpper,
                out russianUpper,
                out error)
                ? englishUpper
                : value;
        }

        private void ApplyKeyboardShortcuts()
        {
            string wheelInput = (_wheelShortcutTextBox.Text ?? string.Empty).Trim();
            string searchInput = (_commandSearchShortcutTextBox.Text ?? string.Empty).Trim();
            string wheelShortcut;
            string searchShortcut;
            string russianUpper;
            string error;

            if (!KeyboardShortcutService.TryNormalizeShortcutInput(
                wheelInput,
                out wheelShortcut,
                out russianUpper,
                out error))
            {
                _wheelShortcutTextBox.Focus();
                MessageBox.Show(
                    this,
                    "Штурвал: " + error,
                    "Горячие клавиши Revit",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!KeyboardShortcutService.TryNormalizeShortcutInput(
                searchInput,
                out searchShortcut,
                out russianUpper,
                out error))
            {
                _commandSearchShortcutTextBox.Focus();
                MessageBox.Show(
                    this,
                    "Команды: " + error,
                    "Горячие клавиши Revit",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            string wheelCommandId = KeyboardShortcutService.FindCommandId(
                _commands,
                typeof(ExternalCommands.CommandsWheel));
            string searchCommandId = KeyboardShortcutService.FindCommandId(
                _commands,
                typeof(ExternalCommands.CommandSearch));

            KeyboardShortcutApplyResult result = KeyboardShortcutService.ApplyToAllInstalledVersions(
                ModuleData.RevitVersion,
                ModuleData.RevitVersionName,
                ModuleData.RibbonTabName,
                wheelShortcut,
                searchShortcut,
                wheelCommandId,
                searchCommandId);

            if (!result.Success)
            {
                _shortcutStatusTextBlock.Text = string.Empty;
                MessageBox.Show(
                    this,
                    result.ErrorMessage,
                    "Горячие клавиши Revit",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _settings.WheelShortcut = wheelShortcut;
            _settings.CommandSearchShortcut = searchShortcut;
            _settings.AreKeyboardShortcutsConfigured = true;
            UserSettingsService.Save(_settings);
            _wheelShortcutTextBox.Text = wheelShortcut;
            _commandSearchShortcutTextBox.Text = searchShortcut;

            string versions = string.Join(
                ", ",
                result.AppliedVersions.Select(version => "Revit " + version).ToArray());
            _shortcutStatusTextBlock.Text = result.Changed
                ? "XML обновлён для: " + versions + "."
                : "В XML уже записаны эти сочетания для: " + versions + ".";
            _shortcutStatusTextBlock.ToolTip = result.FilePath;

            if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
            {
                MessageBox.Show(
                    this,
                    "Часть версий обновить не удалось:\n\n" + result.ErrorMessage,
                    "Горячие клавиши Revit",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void RefreshSettingsControls()
        {
            _isUpdatingSettingsControls = true;

            bool isPinned = string.Equals(_settings.WheelMode, WheelModeNames.Pinned, StringComparison.OrdinalIgnoreCase);
            if (_unpinnedWheelRadioButton != null)
            {
                _unpinnedWheelRadioButton.IsChecked = !isPinned;
            }

            if (_pinnedWheelRadioButton != null)
            {
                _pinnedWheelRadioButton.IsChecked = isPinned;
            }

            if (_wheelCloseButtonCheckBox != null)
            {
                _wheelCloseButtonCheckBox.IsEnabled = !isPinned;
                _wheelCloseButtonCheckBox.IsChecked = isPinned || _settings.IsWheelCloseButtonVisible;
            }

            _isUpdatingSettingsControls = false;
        }

        private void SaveSettingsAndRefresh()
        {
            UserSettingsService.Save(_settings);
            RefreshSettingsControls();
        }

        private TextBox CreateSearchBox()
        {
            TextBox textBox = new TextBox
            {
                FontSize = 16,
                Padding = new Thickness(10, 7, 10, 7),
                Margin = new Thickness(0, 0, 0, 10),
                VerticalContentAlignment = VerticalAlignment.Center
            };

            textBox.TextChanged += delegate { Rebuild(); };

            return textBox;
        }

        private void Rebuild()
        {
            _contentPanel.Children.Clear();

            string query = (_searchBox.Text ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(query))
            {
                RenderSection(
                    "Штурвал",
                    CommandsByIds(_settings.WheelCommandIds),
                    "Добавьте команды зелёным плюсом в строке команды.",
                    CommandListKind.Wheel);

                RenderSection(
                    "Избранное",
                    CommandsByIds(_settings.FavoriteCommandIds),
                    "Добавьте команды сердечком.",
                    CommandListKind.Favorites);

                RenderSection(
                    "Последние",
                    CommandsByIds(_settings.RecentCommandIds),
                    null,
                    CommandListKind.None);

                RenderSection(
                    "Все команды",
                    _commands,
                    null,
                    CommandListKind.None);

                return;
            }

            List<RevitCommandInfo> found = Filter(query).ToList();

            RenderSection(
                string.Format("Найдено: {0}", found.Count),
                found,
                "Ничего не найдено.",
                CommandListKind.None);
        }

        private IEnumerable<RevitCommandInfo> Filter(string query)
        {
            string[] tokens = query.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            return _commands
                .Where(command => tokens.All(token => IsCommandNameMatch(command, token)))
                .OrderByDescending(IsFavorite)
                .ThenByDescending(IsInWheel)
                .ThenBy(command => command.Name);
        }

        private bool IsCommandNameMatch(RevitCommandInfo command, string token)
        {
            if (command == null || string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            return (command.Name ?? string.Empty).IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private IEnumerable<RevitCommandInfo> CommandsByIds(IEnumerable<string> ids)
        {
            if (ids == null)
            {
                yield break;
            }

            foreach (string id in ids)
            {
                RevitCommandInfo command;

                if (!string.IsNullOrWhiteSpace(id) && _commandsById.TryGetValue(id, out command))
                {
                    yield return command;
                }
            }
        }

        private void RenderSection(string title, IEnumerable<RevitCommandInfo> commands, string emptyText, CommandListKind listKind)
        {
            List<RevitCommandInfo> list = commands.ToList();

            if (list.Count == 0 && string.IsNullOrWhiteSpace(emptyText))
            {
                return;
            }

            TextBlock header = new TextBlock
            {
                Text = title,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(70, 70, 70)),
                Margin = new Thickness(2, 12, 0, 6)
            };

            _contentPanel.Children.Add(header);

            if (list.Count == 0)
            {
                _contentPanel.Children.Add(new TextBlock
                {
                    Text = emptyText,
                    Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                    Margin = new Thickness(2, 0, 0, 8)
                });

                return;
            }

            foreach (RevitCommandInfo command in list)
            {
                _contentPanel.Children.Add(CreateCommandRow(command, listKind));
            }
        }






        private UIElement CreateCommandRow(RevitCommandInfo command, CommandListKind listKind)
        {
            Border rowBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 226, 226)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(0, 0, 0, 6),
                Padding = new Thickness(8)
            };

            Grid row = new Grid();

            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Левая зона запуска команды.
            // Теперь команда запускается только при клике по иконке/названию,
            // а не по всей строке до кнопок справа.
            Grid runArea = new Grid
            {
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand
            };

            runArea.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });
            runArea.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Grid.SetColumn(runArea, 0);
            Grid.SetColumnSpan(runArea, 2);
            row.Children.Add(runArea);

            UIElement icon = CreateIcon(command, 26);
            Grid.SetColumn(icon, 0);
            runArea.Children.Add(icon);

            StackPanel textPanel = new StackPanel { Margin = new Thickness(8, 0, 8, 0) };

            textPanel.Children.Add(new TextBlock
            {
                Text = command.Name,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.Black,
                TextWrapping = TextWrapping.Wrap
            });

            textPanel.Children.Add(new TextBlock
            {
                Text = string.Format("{0} / {1}", command.TabName, command.PanelName),
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(110, 110, 110)),
                TextWrapping = TextWrapping.Wrap
            });

            Grid.SetColumn(textPanel, 1);
            runArea.Children.Add(textPanel);

            runArea.MouseLeftButtonUp += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
                Run(command);
            };

            Button favoriteButton = CreateActionButton(IsFavorite(command) ? "\u2665" : "\u2661", "Избранное");
            favoriteButton.Foreground = new SolidColorBrush(Color.FromRgb(190, 45, 70));
            favoriteButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                ToggleFavorite(command);
            };

            Grid.SetColumn(favoriteButton, 2);
            row.Children.Add(favoriteButton);

            bool isInWheel = IsInWheel(command);

            Button wheelButton = CreateActionButton(
                isInWheel ? "\u2212" : "+",
                isInWheel ? "Убрать из штурвала" : "Добавить в штурвал");

            wheelButton.FontSize = 18;
            wheelButton.FontWeight = FontWeights.Bold;

            // Нижний внутренний отступ поднимает плюс/минус чуть выше внутри кнопки.
            wheelButton.Padding = new Thickness(0, 0, 0, 4);

            wheelButton.Foreground = isInWheel
                ? new SolidColorBrush(Color.FromRgb(190, 45, 45))
                : new SolidColorBrush(Color.FromRgb(35, 150, 75));

            wheelButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                ToggleWheel(command);
            };

            Grid.SetColumn(wheelButton, 3);
            row.Children.Add(wheelButton);

            if (listKind != CommandListKind.None)
            {
                if (CanMoveCommand(command, listKind, -1))
                {
                    Button moveUpButton = CreateActionButton(
                        "\u2191",
                        listKind == CommandListKind.Wheel ? "Выше в штурвале" : "Выше в избранном");

                    moveUpButton.Foreground = new SolidColorBrush(Color.FromRgb(72, 72, 72));
                    moveUpButton.Click += delegate (object sender, RoutedEventArgs args)
                    {
                        args.Handled = true;
                        MoveCommand(command, listKind, -1);
                    };

                    Grid.SetColumn(moveUpButton, 4);
                    row.Children.Add(moveUpButton);
                }

                if (CanMoveCommand(command, listKind, 1))
                {
                    Button moveDownButton = CreateActionButton(
                        "\u2193",
                        listKind == CommandListKind.Wheel ? "Ниже в штурвале" : "Ниже в избранном");

                    moveDownButton.Foreground = new SolidColorBrush(Color.FromRgb(72, 72, 72));
                    moveDownButton.Click += delegate (object sender, RoutedEventArgs args)
                    {
                        args.Handled = true;
                        MoveCommand(command, listKind, 1);
                    };

                    Grid.SetColumn(moveDownButton, 5);
                    row.Children.Add(moveDownButton);
                }
            }

            rowBorder.Child = row;

            return rowBorder;
        }

        private UIElement CreateIcon(RevitCommandInfo command, double size)
        {
            ImageSource source = IconSourceLoader.Load(command);

            if (source != null)
            {
                return new Image
                {
                    Source = source,
                    Width = size,
                    Height = size,
                    Stretch = Stretch.Uniform,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };
            }

            string letter = string.IsNullOrWhiteSpace(command.Name)
                ? "?"
                : command.Name.Substring(0, 1).ToUpperInvariant();

            return new Border
            {
                Width = size + 2,
                Height = size + 2,
                CornerRadius = new CornerRadius(5),
                Background = new SolidColorBrush(Color.FromRgb(230, 235, 240)),
                Child = new TextBlock
                {
                    Text = letter,
                    FontWeight = FontWeights.SemiBold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
        }

        private Button CreateActionButton(string text, string tooltip)
        {
            return new Button
            {
                Content = text,
                ToolTip = tooltip,
                Width = 32,
                Height = 30,
                Margin = new Thickness(3, 0, 0, 0),
                FontSize = 15,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(214, 214, 214)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center
            };
        }

        private void ToggleFavorite(RevitCommandInfo command)
        {
            if (IsFavorite(command))
            {
                _settings.FavoriteCommandIds.RemoveAll(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                _settings.FavoriteCommandIds.Insert(0, command.Id);
            }

            UserSettingsService.Save(_settings);
            Rebuild();
        }

        private void ToggleWheel(RevitCommandInfo command)
        {
            if (IsInWheel(command))
            {
                _settings.WheelCommandIds.RemoveAll(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                if (_settings.WheelCommandIds.Count >= 8)
                {
                    MessageBox.Show(
                        this,
                        "В штурвал можно добавить не больше 8 команд.",
                        "Штурвал",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    return;
                }

                _settings.WheelCommandIds.Add(command.Id);
            }

            UserSettingsService.Save(_settings);
            Rebuild();
        }

        private bool CanMoveCommand(RevitCommandInfo command, CommandListKind listKind, int direction)
        {
            List<string> ids = GetCommandIdList(listKind);

            if (ids == null)
            {
                return false;
            }

            int index = GetCommandIndex(command, ids);

            if (index < 0)
            {
                return false;
            }

            int targetIndex = index + direction;

            return targetIndex >= 0 && targetIndex < ids.Count;
        }

        private void MoveCommand(RevitCommandInfo command, CommandListKind listKind, int direction)
        {
            List<string> ids = GetCommandIdList(listKind);

            if (ids == null)
            {
                return;
            }

            int index = GetCommandIndex(command, ids);
            int targetIndex = index + direction;

            if (index < 0 || targetIndex < 0 || targetIndex >= ids.Count)
            {
                return;
            }

            string currentId = ids[index];
            ids[index] = ids[targetIndex];
            ids[targetIndex] = currentId;

            UserSettingsService.Save(_settings);
            Rebuild();
        }

        private int GetCommandIndex(RevitCommandInfo command, List<string> ids)
        {
            if (command == null || ids == null || string.IsNullOrWhiteSpace(command.Id))
                return -1;

            return ids.FindIndex(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
        }

        private List<string> GetCommandIdList(CommandListKind listKind)
        {
            switch (listKind)
            {
                case CommandListKind.Wheel:
                    return _settings.WheelCommandIds;

                case CommandListKind.Favorites:
                    return _settings.FavoriteCommandIds;

                default:
                    return null;
            }
        }

        private void Run(RevitCommandInfo command)
        {
            UserSettingsService.AddRecent(_settings, command.Id);
            UserSettingsService.Save(_settings);
            Rebuild();
            _executor.Run(command);
        }

        private bool IsFavorite(RevitCommandInfo command)
        {
            return command != null
                && _settings.FavoriteCommandIds.Any(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsInWheel(RevitCommandInfo command)
        {
            return command != null
                && _settings.WheelCommandIds.Any(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
        }

        private static T FindParent<T>(DependencyObject source) where T : DependencyObject
        {
            DependencyObject current = source;

            while (current != null)
            {
                T typed = current as T;

                if (typed != null)
                {
                    return typed;
                }

                FrameworkContentElement contentElement = current as FrameworkContentElement;
                current = contentElement != null ? contentElement.Parent : VisualTreeHelper.GetParent(current);
            }

            return null;
        }
    }
}
