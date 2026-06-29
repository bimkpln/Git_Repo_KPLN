using KPLN_CommandsWheel.ExternalCommands;
using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using KPLN_Library_PluginActivityWorker;
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
        private readonly List<RevitCommandInfo> _commands;
        private readonly Dictionary<string, RevitCommandInfo> _commandsById;
        private readonly UserSettings _settings;
        private readonly RevitCommandExecutor _executor;
        private readonly TextBox _searchBox;
        private readonly StackPanel _contentPanel;

        internal CommandSearchWindow(IEnumerable<RevitCommandInfo> commands, UserSettings settings, RevitCommandExecutor executor)
        {
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

            Rebuild();
        }

        private UIElement CreateContent()
        {
            Grid root = new Grid { Margin = new Thickness(14) };
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
                RenderSection("Штурвал", CommandsByIds(_settings.WheelCommandIds), "Добавьте команды кнопкой " + "\u2638" + " в строке команды.");
                RenderSection("Избранное", CommandsByIds(_settings.FavoriteCommandIds), "Добавьте команды сердечком.");
                RenderSection("Последние", CommandsByIds(_settings.RecentCommandIds), null);
                RenderSection("Все команды", _commands, null);
                return;
            }

            List<RevitCommandInfo> found = Filter(query).ToList();
            RenderSection(string.Format("Найдено: {0}", found.Count), found, "Ничего не найдено.");
        }

        private IEnumerable<RevitCommandInfo> Filter(string query)
        {
            string[] tokens = query.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return _commands
                .Where(command => tokens.All(token => (command.SearchText ?? string.Empty).IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0))
                .OrderByDescending(IsFavorite)
                .ThenByDescending(IsInWheel)
                .ThenBy(command => command.Name);
        }

        private IEnumerable<RevitCommandInfo> CommandsByIds(IEnumerable<string> ids)
        {
            if (ids == null)
                yield break;

            foreach (string id in ids)
            {
                if (!string.IsNullOrWhiteSpace(id) && _commandsById.TryGetValue(id, out RevitCommandInfo command))
                    yield return command;
            }
        }

        private void RenderSection(string title, IEnumerable<RevitCommandInfo> commands, string emptyText)
        {
            List<RevitCommandInfo> list = commands.ToList();
            if (list.Count == 0 && string.IsNullOrWhiteSpace(emptyText))
                return;

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
                _contentPanel.Children.Add(CreateCommandRow(command));
            }
        }

        private UIElement CreateCommandRow(RevitCommandInfo command)
        {
            Border rowBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 226, 226)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(0, 0, 0, 6),
                Padding = new Thickness(8),
                Cursor = Cursors.Hand
            };

            Grid row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            UIElement icon = CreateIcon(command, 26);
            Grid.SetColumn(icon, 0);
            row.Children.Add(icon);

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
            row.Children.Add(textPanel);

            Button favoriteButton = CreateActionButton(IsFavorite(command) ? "\u2665" : "\u2661", "Избранное");
            favoriteButton.Foreground = new SolidColorBrush(Color.FromRgb(190, 45, 70));
            favoriteButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                ToggleFavorite(command);
            };
            Grid.SetColumn(favoriteButton, 2);
            row.Children.Add(favoriteButton);

            Button wheelButton = CreateActionButton("\u2638", IsInWheel(command) ? "Убрать из штурвала" : "Добавить в штурвал");
            wheelButton.Foreground = IsInWheel(command)
                ? new SolidColorBrush(Color.FromRgb(26, 110, 170))
                : new SolidColorBrush(Color.FromRgb(190, 190, 190));
            wheelButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                ToggleWheel(command);
            };
            Grid.SetColumn(wheelButton, 3);
            row.Children.Add(wheelButton);

            Button moveUpButton = CreateActionButton("\u2191", "Выше в штурвале");
            moveUpButton.IsEnabled = CanMoveWheelCommand(command, -1);
            moveUpButton.Foreground = new SolidColorBrush(Color.FromRgb(72, 72, 72));
            moveUpButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                MoveWheelCommand(command, -1);
            };
            Grid.SetColumn(moveUpButton, 4);
            row.Children.Add(moveUpButton);

            Button moveDownButton = CreateActionButton("\u2193", "Ниже в штурвале");
            moveDownButton.IsEnabled = CanMoveWheelCommand(command, 1);
            moveDownButton.Foreground = new SolidColorBrush(Color.FromRgb(72, 72, 72));
            moveDownButton.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                MoveWheelCommand(command, 1);
            };
            Grid.SetColumn(moveDownButton, 5);
            row.Children.Add(moveDownButton);

            rowBorder.Child = row;
            rowBorder.MouseLeftButtonUp += delegate (object sender, MouseButtonEventArgs args)
            {
                if (FindParent<Button>(args.OriginalSource as DependencyObject) != null)
                    return;

                Run(command);
            };

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

            string letter = string.IsNullOrWhiteSpace(command.Name) ? "?" : command.Name.Substring(0, 1).ToUpperInvariant();
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
                    MessageBox.Show(this, "В штурвал можно добавить не больше 8 команд.", "Штурвал", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _settings.WheelCommandIds.Add(command.Id);
            }

            UserSettingsService.Save(_settings);
            Rebuild();
        }

        private bool CanMoveWheelCommand(RevitCommandInfo command, int direction)
        {
            int index = GetWheelCommandIndex(command);
            if (index < 0)
            {
                return false;
            }

            int targetIndex = index + direction;
            return targetIndex >= 0 && targetIndex < _settings.WheelCommandIds.Count;
        }

        private void MoveWheelCommand(RevitCommandInfo command, int direction)
        {
            int index = GetWheelCommandIndex(command);
            int targetIndex = index + direction;
            if (index < 0 || targetIndex < 0 || targetIndex >= _settings.WheelCommandIds.Count)
            {
                return;
            }

            string currentId = _settings.WheelCommandIds[index];
            _settings.WheelCommandIds[index] = _settings.WheelCommandIds[targetIndex];
            _settings.WheelCommandIds[targetIndex] = currentId;

            UserSettingsService.Save(_settings);
            Rebuild();
        }

        private int GetWheelCommandIndex(RevitCommandInfo command)
        {
            if (command == null || string.IsNullOrWhiteSpace(command.Id))
                return -1;

            return _settings.WheelCommandIds.FindIndex(id => string.Equals(id, command.Id, StringComparison.OrdinalIgnoreCase));
        }

        private void Run(RevitCommandInfo command)
        {
            UserSettingsService.AddRecent(_settings, command.Id);
            UserSettingsService.Save(_settings);
            Rebuild();
            _executor.Run(command);

            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(CommandSearch.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
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