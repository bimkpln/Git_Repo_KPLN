using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

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
        private readonly HashSet<string> _capturedHotkeyKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _pressedHotkeyKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private RadioButton _unpinnedWheelRadioButton;
        private RadioButton _pinnedWheelRadioButton;
        private CheckBox _wheelCloseButtonCheckBox;
        private Button _commandSearchHotkeyButton;
        private Button _commandsWheelHotkeyButton;
        private bool _isUpdatingSettingsControls;
        private bool _isCapturingHotkey;
        private HotkeyTarget _capturingHotkeyTarget;

        private enum CommandListKind
        {
            None,
            Wheel,
            Favorites
        }

        private enum HotkeyTarget
        {
            CommandSearch,
            CommandsWheel
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
                if (_isCapturingHotkey)
                {
                    args.Handled = true;
                    CaptureKeyboardHotkey(args);
                    return;
                }

                if (args.Key == Key.Escape)
                {
                    args.Handled = true;
                    Close();
                }
            };
            PreviewKeyUp += delegate (object sender, KeyEventArgs args)
            {
                if (_isCapturingHotkey)
                {
                    args.Handled = true;
                    ReleaseKeyboardHotkey(args);
                }
            };
            PreviewMouseDown += delegate (object sender, MouseButtonEventArgs args)
            {
                if (_isCapturingHotkey)
                {
                    args.Handled = true;
                    CaptureMouseHotkey(args);
                }
            };
            Closed += delegate
            {
                if (ReferenceEquals(_current, this))
                {
                    _current = null;
                }

                if (_isCapturingHotkey)
                {
                    _isCapturingHotkey = false;
                    _capturedHotkeyKeys.Clear();
                    _pressedHotkeyKeys.Clear();
                    HotkeyService.ResumeHotkeys();
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

            panel.Children.Add(CreateSettingsHeader("Горячие клавиши"));
            panel.Children.Add(CreateHotkeyRow("Окно Команды", HotkeyTarget.CommandSearch, out _commandSearchHotkeyButton));
            panel.Children.Add(CreateHotkeyRow("Штурвал", HotkeyTarget.CommandsWheel, out _commandsWheelHotkeyButton));
            panel.Children.Add(CreateHotkeyHelp());

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

        private UIElement CreateHotkeyHelp()
        {
            Border border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 226, 226)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 10, 0, 0)
            };

            StackPanel stackPanel = new StackPanel();

            TextBlock text = new TextBlock
            {
                Text = "Можно назначить от одной до трёх любых клавиш клавиатуры.\nДля мыши доступны только боковые кнопки XButton1 и XButton2.\nВажно: ЛКМ, ПКМ и колесо не назначаются.\nИзменения горячих клавиш применяются после перезапуска Revit.",
                Foreground = new SolidColorBrush(Color.FromRgb(92, 92, 92)),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 18,
                Margin = new Thickness(0, 0, 0, 10)
            };
            stackPanel.Children.Add(text);

            stackPanel.Children.Add(CreateMouseButtonsImage());

            border.Child = stackPanel;
            return border;
        }

        private UIElement CreateMouseButtonsImage()
        {
            const string resourceName = "KPLN_CommandsWheel.Imagens.mouseSideButtons.png";

            Stream stream = typeof(CommandSearchWindow).Assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                return new TextBlock
                {
                    Text = "XButton1 — верхняя боковая кнопка, XButton2 — нижняя боковая кнопка.",
                    Foreground = new SolidColorBrush(Color.FromRgb(92, 92, 92)),
                    TextWrapping = TextWrapping.Wrap
                };
            }

            using (stream)
            {
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();

                return new Image
                {
                    Source = image,
                    Width = 200,
                    Stretch = Stretch.Uniform,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 2, 0, 0)
                };
            }
        }

        private UIElement CreateHotkeyRow(string title, HotkeyTarget target, out Button hotkeyButton)
        {
            Grid row = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            row.Children.Add(new TextBlock
            {
                Text = title,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Black
            });

            hotkeyButton = new Button
            {
                Height = 30,
                Margin = new Thickness(0, 0, 6, 0),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(10, 0, 10, 0)
            };
            hotkeyButton.Click += delegate
            {
                StartHotkeyCapture(target);
            };
            Grid.SetColumn(hotkeyButton, 1);
            row.Children.Add(hotkeyButton);

            Button clearButton = new Button
            {
                Content = "Очистить",
                Height = 30,
                MinWidth = 80,
                Padding = new Thickness(10, 0, 10, 0)
            };
            clearButton.Click += delegate
            {
                ClearHotkey(target);
            };
            Grid.SetColumn(clearButton, 2);
            row.Children.Add(clearButton);

            return row;
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

            if (_commandSearchHotkeyButton != null)
            {
                _commandSearchHotkeyButton.Content = GetHotkeyButtonText(HotkeyTarget.CommandSearch);
            }

            if (_commandsWheelHotkeyButton != null)
            {
                _commandsWheelHotkeyButton.Content = GetHotkeyButtonText(HotkeyTarget.CommandsWheel);
            }

            _isUpdatingSettingsControls = false;
        }

        private string GetHotkeyButtonText(HotkeyTarget target)
        {
            if (_isCapturingHotkey && _capturingHotkeyTarget == target)
            {
                return "Нажмите сочетание...";
            }

            return HotkeyGestureService.ToDisplayText(GetHotkey(target));
        }

        private void StartHotkeyCapture(HotkeyTarget target)
        {
            _isCapturingHotkey = true;
            _capturingHotkeyTarget = target;
            _capturedHotkeyKeys.Clear();
            _pressedHotkeyKeys.Clear();
            HotkeyService.SuspendHotkeys();
            RefreshSettingsControls();
            Dispatcher.BeginInvoke(new Action(delegate
            {
                Focus();
                Keyboard.Focus(this);
            }), DispatcherPriority.Input);
        }

        private void StopHotkeyCapture()
        {
            _isCapturingHotkey = false;
            _capturedHotkeyKeys.Clear();
            _pressedHotkeyKeys.Clear();
            HotkeyService.ResumeHotkeys();
            RefreshSettingsControls();
            _searchBox.Focus();
            Keyboard.Focus(_searchBox);
        }

        private void CaptureKeyboardHotkey(KeyEventArgs args)
        {
            if (args.Key == Key.Escape)
            {
                StopHotkeyCapture();
                return;
            }

            string keyName = HotkeyGestureService.GetKeyName(args);
            if (string.IsNullOrWhiteSpace(keyName) || args.IsRepeat)
            {
                return;
            }

            _capturedHotkeyKeys.Add(keyName);
            _pressedHotkeyKeys.Add(keyName);

            if (_capturedHotkeyKeys.Count > 3)
            {
                MessageBox.Show(this, "В сочетании может быть не больше трёх клавиш клавиатуры.", "Горячие клавиши", MessageBoxButton.OK, MessageBoxImage.Information);
                StopHotkeyCapture();
            }
        }

        private void ReleaseKeyboardHotkey(KeyEventArgs args)
        {
            string keyName = HotkeyGestureService.GetKeyName(args);
            if (!string.IsNullOrWhiteSpace(keyName))
            {
                _pressedHotkeyKeys.Remove(keyName);
            }

            if (_capturedHotkeyKeys.Count == 0 || _pressedHotkeyKeys.Count != 0)
            {
                return;
            }

            AssignHotkey(_capturingHotkeyTarget, new HotkeyGesture
            {
                Keys = HotkeyGestureService.NormalizeKeys(_capturedHotkeyKeys)
            });
        }

        private void CaptureMouseHotkey(MouseButtonEventArgs args)
        {
            HotkeyGesture gesture = HotkeyGestureService.FromMouseEvent(args, _pressedHotkeyKeys);
            if (HotkeyGestureService.IsEmpty(gesture))
            {
                return;
            }

            AssignHotkey(_capturingHotkeyTarget, gesture);
        }

        private void AssignHotkey(HotkeyTarget target, HotkeyGesture gesture)
        {
            if (gesture.Keys != null && gesture.Keys.Count > 3)
            {
                MessageBox.Show(this, "В сочетании может быть не больше трёх клавиш клавиатуры.", "Горячие клавиши", MessageBoxButton.OK, MessageBoxImage.Information);
                StopHotkeyCapture();
                return;
            }

            HotkeyTarget otherTarget = target == HotkeyTarget.CommandSearch
                ? HotkeyTarget.CommandsWheel
                : HotkeyTarget.CommandSearch;

            if (!HotkeyGestureService.IsEmpty(gesture) && HotkeyGestureService.AreEqual(gesture, GetHotkey(otherTarget)))
            {
                MessageBox.Show(this, "Для окна Команды и Штурвала нужны разные сочетания.", "Горячие клавиши", MessageBoxButton.OK, MessageBoxImage.Information);
                StopHotkeyCapture();
                return;
            }

            SetHotkey(target, gesture);
            SaveSettingsAndRefresh();
            StopHotkeyCapture();
        }

        private void ClearHotkey(HotkeyTarget target)
        {
            SetHotkey(target, new HotkeyGesture());
            SaveSettingsAndRefresh();
        }

        private HotkeyGesture GetHotkey(HotkeyTarget target)
        {
            return target == HotkeyTarget.CommandSearch
                ? _settings.CommandSearchHotkey
                : _settings.CommandsWheelHotkey;
        }

        private void SetHotkey(HotkeyTarget target, HotkeyGesture gesture)
        {
            if (target == HotkeyTarget.CommandSearch)
            {
                _settings.CommandSearchHotkey = gesture;
                return;
            }

            _settings.CommandsWheelHotkey = gesture;
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