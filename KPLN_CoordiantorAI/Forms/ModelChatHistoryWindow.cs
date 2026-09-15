using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using KPLN_CoordiantorAI.ExternalModel;

namespace KPLN_CoordiantorAI.Forms
{
    internal sealed class ModelChatHistoryWindow : Window
    {
        private const int PageSize = 10;
        private readonly ModelChatHistoryRepository _repository;
        private readonly ModelChatLegacyLogReader _legacyLogReader = new ModelChatLegacyLogReader();
        private readonly string _legacyLogFolder;
        private readonly List<ModelChatHistoryEntry> _entries = new List<ModelChatHistoryEntry>();
        private readonly Dictionary<ModelChatHistoryEntry, FrameworkElement> _entryBlocks =
            new Dictionary<ModelChatHistoryEntry, FrameworkElement>();
        private readonly StackPanel _contentPanel;
        private readonly Button _loadMoreButton;
        private readonly TextBlock _statusText;
        private readonly ScrollViewer _scrollViewer;
        private int _databaseOffset;
        private int _legacyOffset;
        private bool _databaseExhausted;
        private IList<ModelChatHistoryEntry> _legacyEntries;
        private bool _isLoading;

        public ModelChatHistoryWindow(ModelChatHistoryRepository repository, string legacyLogFolder)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _legacyLogFolder = legacyLogFolder;
            Title = "История чата";
            Width = 820;
            Height = 650;
            MinWidth = 620;
            MinHeight = 440;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = Brush(32, 36, 45);
            FontFamily = new FontFamily("Segoe UI");

            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Border header = new Border
            {
                Background = Brush(38, 43, 53),
                BorderBrush = Brush(52, 59, 73),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(16, 12, 16, 12),
                Child = new TextBlock
                {
                    Text = "История чата",
                    FontSize = 17,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.White
                }
            };
            root.Children.Add(header);

            _contentPanel = new StackPanel { Margin = new Thickness(12, 6, 12, 12) };
            _scrollViewer = new ScrollViewer
            {
                Content = _contentPanel,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            Grid.SetRow(_scrollViewer, 1);
            root.Children.Add(_scrollViewer);

            Grid footer = new Grid { Margin = new Thickness(12, 0, 12, 12) };
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _statusText = new TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brush(168, 176, 190)
            };
            footer.Children.Add(_statusText);
            _loadMoreButton = new Button
            {
                Content = "Загрузить ещё 10",
                MinHeight = 34,
                Padding = new Thickness(14, 6, 14, 6),
                Background = Brush(42, 48, 59),
                Foreground = Brushes.White,
                BorderBrush = Brush(75, 85, 104),
                BorderThickness = new Thickness(1),
                Visibility = Visibility.Collapsed
            };
            _loadMoreButton.Click += LoadMoreButton_Click;
            Grid.SetColumn(_loadMoreButton, 1);
            footer.Children.Add(_loadMoreButton);
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            Content = root;
            Loaded += OnLoaded;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= OnLoaded;
            await LoadNextPageAsync();
        }

        private async void LoadMoreButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadNextPageAsync();
        }

        private async Task LoadNextPageAsync()
        {
            if (_isLoading)
                return;

            _isLoading = true;
            _loadMoreButton.IsEnabled = false;
            _statusText.Text = "Загружаю историю...";
            try
            {
                HistoryPage page = await Task.Run(LoadNextPageFromSources);
                ScrollAnchor anchor = CaptureScrollAnchor();
                _entries.AddRange(page.Entries);
                RenderEntries();
                _scrollViewer.UpdateLayout();
                if (anchor == null)
                    _scrollViewer.ScrollToEnd();
                else
                    RestoreScrollAnchor(anchor);
                _loadMoreButton.Visibility = page.HasMore ? Visibility.Visible : Visibility.Collapsed;
                _statusText.Text = !string.IsNullOrWhiteSpace(page.Warning)
                    ? "Показано запросов: " + _entries.Count.ToString(CultureInfo.InvariantCulture)
                        + ". TXT-история недоступна: " + page.Warning
                    : (_entries.Count == 0
                    ? "Сохранённых вопросов и ответов пока нет."
                    : "Показано запросов: " + _entries.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                _statusText.Text = "История недоступна: " + ex.Message;
                _loadMoreButton.Visibility = Visibility.Collapsed;
            }
            finally
            {
                _isLoading = false;
                _loadMoreButton.IsEnabled = true;
            }
        }

        private HistoryPage LoadNextPageFromSources()
        {
            List<ModelChatHistoryEntry> pageEntries = new List<ModelChatHistoryEntry>();
            bool databaseHasMore = false;
            if (!_databaseExhausted)
            {
                IList<ModelChatHistoryEntry> databasePage =
                    _repository.GetCompletedHistoryPage(_databaseOffset, PageSize, out databaseHasMore);
                pageEntries.AddRange(databasePage);
                _databaseOffset += databasePage.Count;
                _databaseExhausted = !databaseHasMore;
            }

            if (!_databaseExhausted)
            {
                return new HistoryPage
                {
                    Entries = pageEntries,
                    HasMore = true
                };
            }

            string warning = null;
            if (_legacyEntries == null)
            {
                try
                {
                    _legacyEntries = _legacyLogReader.ReadMissingEntries(
                        _legacyLogFolder,
                        _entries.Concat(pageEntries));
                }
                catch (Exception ex)
                {
                    _legacyEntries = new List<ModelChatHistoryEntry>();
                    warning = ex.Message;
                }
            }

            int remainingCapacity = PageSize - pageEntries.Count;
            if (remainingCapacity > 0 && _legacyOffset < _legacyEntries.Count)
            {
                List<ModelChatHistoryEntry> legacyPage = _legacyEntries
                    .Skip(_legacyOffset)
                    .Take(remainingCapacity)
                    .ToList();
                pageEntries.AddRange(legacyPage);
                _legacyOffset += legacyPage.Count;
            }

            return new HistoryPage
            {
                Entries = pageEntries,
                HasMore = _legacyOffset < _legacyEntries.Count,
                Warning = warning
            };
        }

        private void RenderEntries()
        {
            _contentPanel.Children.Clear();
            _entryBlocks.Clear();
            List<IGrouping<string, ModelChatHistoryEntry>> groups = _entries
                .GroupBy(entry => (entry.ModelIdentity ?? string.Empty) + "|" + entry.RevitVersion)
                .OrderBy(group => group.Max(entry => ParseDate(entry.RequestTime)))
                .ToList();

            foreach (IGrouping<string, ModelChatHistoryEntry> group in groups)
            {
                ModelChatHistoryEntry newest = group.OrderByDescending(entry => ParseDate(entry.RequestTime)).First();
                _contentPanel.Children.Add(CreateGroupHeader(newest));
                foreach (ModelChatHistoryEntry entry in group.OrderBy(item => ParseDate(item.RequestTime)))
                {
                    FrameworkElement block = CreateRequestBlock(entry);
                    _entryBlocks.Add(entry, block);
                    _contentPanel.Children.Add(block);
                }
            }
        }

        private ScrollAnchor CaptureScrollAnchor()
        {
            if (_entryBlocks.Count == 0)
                return null;

            ScrollAnchor topmostVisible = null;
            foreach (KeyValuePair<ModelChatHistoryEntry, FrameworkElement> item in _entryBlocks)
            {
                double top = item.Value.TransformToAncestor(_scrollViewer).Transform(new Point(0, 0)).Y;
                if (top + item.Value.ActualHeight > 0
                    && top < _scrollViewer.ViewportHeight
                    && (topmostVisible == null || top < topmostVisible.Top))
                {
                    topmostVisible = new ScrollAnchor
                    {
                        Entry = item.Key,
                        Top = top,
                        VerticalOffset = _scrollViewer.VerticalOffset
                    };
                }
            }

            if (topmostVisible != null)
                return topmostVisible;

            return new ScrollAnchor { VerticalOffset = _scrollViewer.VerticalOffset };
        }

        private void RestoreScrollAnchor(ScrollAnchor anchor)
        {
            FrameworkElement block;
            if (anchor.Entry == null || !_entryBlocks.TryGetValue(anchor.Entry, out block))
            {
                _scrollViewer.ScrollToVerticalOffset(anchor.VerticalOffset);
                return;
            }

            double newTop = block.TransformToAncestor(_scrollViewer).Transform(new Point(0, 0)).Y;
            _scrollViewer.ScrollToVerticalOffset(
                _scrollViewer.VerticalOffset + newTop - anchor.Top);
        }

        private static FrameworkElement CreateGroupHeader(ModelChatHistoryEntry entry)
        {
            string modelName = string.IsNullOrWhiteSpace(entry.ModelName) ? "Без имени" : entry.ModelName.Trim();
            string versionLabel = entry.RevitVersion > 0
                ? "Revit " + entry.RevitVersion.ToString(CultureInfo.InvariantCulture)
                : "версия Revit не определена";
            return new Border
            {
                Margin = new Thickness(0, 12, 0, 8),
                Padding = new Thickness(0, 0, 0, 7),
                BorderBrush = Brush(75, 85, 104),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Child = new TextBlock
                {
                    Text = modelName + " - " + versionLabel,
                    FontSize = 15,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.White,
                    TextWrapping = TextWrapping.Wrap
                }
            };
        }

        private static FrameworkElement CreateRequestBlock(ModelChatHistoryEntry entry)
        {
            StackPanel block = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
            DateTime requestTime = ParseDate(entry.RequestTime);
            block.Children.Add(new TextBlock
            {
                Text = requestTime == DateTime.MinValue ? entry.RequestTime : requestTime.ToString("dd.MM.yyyy HH:mm"),
                FontSize = 11,
                Foreground = Brush(138, 148, 165),
                Margin = new Thickness(2, 0, 2, 4)
            });
            block.Children.Add(CreateMessage("Вы", entry.UserQuestion, Brush(44, 107, 237), HorizontalAlignment.Right));
            block.Children.Add(CreateMessage("ИИ", entry.FinalAnswer, Brush(48, 54, 66), HorizontalAlignment.Left));
            return block;
        }

        private static FrameworkElement CreateMessage(
            string author,
            string text,
            Brush background,
            HorizontalAlignment alignment)
        {
            StackPanel content = new StackPanel();
            content.Children.Add(new TextBlock
            {
                Text = author,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brush(205, 211, 221),
                Margin = new Thickness(0, 0, 0, 3)
            });
            content.Children.Add(new TextBlock
            {
                Text = text ?? string.Empty,
                FontSize = 13,
                Foreground = Brushes.White,
                TextWrapping = TextWrapping.Wrap
            });

            return new Border
            {
                Child = content,
                Background = background,
                CornerRadius = new CornerRadius(7),
                Padding = new Thickness(11, 8, 11, 8),
                Margin = new Thickness(0, 2, 0, 5),
                HorizontalAlignment = alignment,
                MaxWidth = 680
            };
        }

        private static DateTime ParseDate(string value)
        {
            DateTime result;
            if (DateTime.TryParseExact(
                value,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out result))
                return result;
            return DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out result)
                ? result
                : DateTime.MinValue;
        }

        private static Brush Brush(byte red, byte green, byte blue)
        {
            return new SolidColorBrush(Color.FromRgb(red, green, blue));
        }

        private sealed class HistoryPage
        {
            public IList<ModelChatHistoryEntry> Entries { get; set; }
            public bool HasMore { get; set; }
            public string Warning { get; set; }
        }

        private sealed class ScrollAnchor
        {
            public ModelChatHistoryEntry Entry { get; set; }
            public double Top { get; set; }
            public double VerticalOffset { get; set; }
        }
    }
}
