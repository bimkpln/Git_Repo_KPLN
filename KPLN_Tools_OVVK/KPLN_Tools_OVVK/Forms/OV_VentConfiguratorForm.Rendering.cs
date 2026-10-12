using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OV_VentConfiguratorForm : Window
    {
        private readonly Dictionary<string, FrameworkElement> _parameterEditors = new Dictionary<string, FrameworkElement>(StringComparer.Ordinal);
        private TextBlock _installationLengthText;
        private WrapPanel _sectionParametersPanel;
        private Point _dragStart;
        private static readonly string[] SectionIconNames =
        {
            "Вентилятор",
            "Водяной нагреватель",
            "Водяной охладитель",
            "Воздушный клапан",
            "Гибкая вставка",
            "Фильтр",
            "Шумоглушитель"
        };
        private readonly Dictionary<string, ImageSource> _sectionIcons = new Dictionary<string, ImageSource>(StringComparer.Ordinal);

        internal void NavigateToParameter(string name)
        {
            if (Model.Configuration == null)
                return;
            name = FamilyParameterNames.Canonical(name);
            if (Model.Configuration.Info.Assignments().ContainsKey(name))
            {
                ConfigurationTabs.SelectedIndex = 0;
                var names = new[]
                {
                    InstallationInfo.SystemNameParameter,
                    InstallationInfo.ManufacturerParameter,
                    InstallationInfo.MarkParameter,
                    InstallationInfo.UnitParameter,
                    InstallationInfo.DescriptionParameter,
                    InstallationInfo.ProductCodeParameter,
                    InstallationInfo.MassTextParameter
                };
                var box = InfoPanel.Children.OfType<TextBox>().ElementAtOrDefault(Array.IndexOf(names, name));
                if (box != null)
                {
                    box.BringIntoView();
                    box.Focus();
                    box.SelectAll();
                }

                return;
            }

            ConfigurationTabs.SelectedIndex = 1;
            for (int i = 0; i < Model.Configuration.Blocks.Count; i++)
            {
                int slot = Model.Configuration.SlotAt(i);
                string prefix = InstallationConfiguration.ParameterName(slot);
                prefix = prefix.Substring(0, prefix.Length - "Тип".Length);
                if (!name.StartsWith(prefix, StringComparison.Ordinal) && !SharedParameters.ForSection(Model.Configuration.Blocks[i].Type).Contains(name))
                    continue;
                CurrentOV_VentConfiguratorVM.SelectSection(Model.Configuration.Blocks[i]);
                break;
            }

            FrameworkElement editor;
            if (name == "Секции_Промежуточные_Количество")
                editor = SectionCountComboBox;
            else if (name == ValveControl.ParameterName)
                editor = AddValveCheckBox;
            else
                _parameterEditors.TryGetValue(name, out editor);
            if (editor != null)
            {
                editor.BringIntoView();
                editor.Focus();
                (editor as TextBox)?.SelectAll();
            }
        }

        internal void RenderConfiguration(InstallationConfiguration configuration)
        {
            _parameterEditors.Clear();
            InfoPanel.DataContext = configuration.Info;
            var overall = new UniformGrid
            {
                Columns = 3
            };
            var width = DimensionEditor("Ширина установки", configuration.InstallationWidth, "Установка_Ширина");
            var height = DimensionEditor("Высота установки", configuration.InstallationHeight, "Установка_Высота");
            width.Margin = height.Margin = new Thickness(0, 0, 16, 0);
            overall.Children.Add(width);
            overall.Children.Add(height);
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
            length.Children.Add(_installationLengthText);
            overall.Children.Add(length);
            InstallationWidthEditor.Content = overall;
            var additional = new UniformGrid
            {
                Columns = 4
            };
            additional.Children.Add(BooleanEditor("Рама", configuration.SharedBooleans[SharedParameters.FrameVisible], SharedParameters.FrameVisible, false));
            additional.Children.Add(DimensionEditor("Высота рамы", configuration.FrameHeight, "Основание_Рама_Высота"));
            additional.Children.Add(BooleanEditor("Зона обслуживания", configuration.SharedBooleans[SharedParameters.ServiceRight], SharedParameters.ServiceRight, true));
            additional.Children.Add(DimensionEditor("Зона обслуживания: глубина", configuration.SharedDimensions[SharedParameters.ServiceDepth], SharedParameters.ServiceDepth));
            foreach (FrameworkElement field in additional.Children)
                field.Margin = new Thickness(0, 0, 16, 0);
            AdditionalInstallationEditor.Content = additional;
        }

        private static Binding ValueBinding(object source, string property)
        {
            return new Binding(property)
            {
                Source = source,
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            };
        }

        internal void SyncSectionControls()
        {
            bool previous = Model.IsSynchronizing;
            Model.IsSynchronizing = true;
            try
            {
                SectionCountComboBox.ItemsSource = Enumerable.Range(Model.Configuration.Minimum, Model.Configuration.Capacity - Model.Configuration.Minimum + 1).ToList();
                SectionCountComboBox.SelectedItem = Model.Configuration.SectionCount;
                SectionCountComboBox.ToolTip = ParameterTooltip("Секции_Промежуточные_Количество", nameOnly: true);
                AddValveCheckBox.IsChecked = Model.Configuration.HasValve;
                AddValveCheckBox.IsEnabled = Model.Configuration.CanChangeValve;
                AddValveCheckBox.ToolTip = null;
            }
            finally
            {
                Model.IsSynchronizing = previous;
            }
        }

        private static string SectionIconFileName(SectionTypeChoice type)
        {
            if (type == null || type.IsEmpty)
                return null;
            // Вид оборудования определяет иконку независимо от позиции секции/соединителя.
            foreach (string name in SectionIconNames)
                if ((type.FamilyName ?? string.Empty).IndexOf("_Секция_" + name + "_", StringComparison.OrdinalIgnoreCase) >= 0 || string.Equals(type.FamilyName, name, StringComparison.OrdinalIgnoreCase))
                    return name + ".png";
            foreach (string name in SectionIconNames)
                if (string.Equals(type.TypeName, name, StringComparison.OrdinalIgnoreCase) || (type.TypeName ?? string.Empty).StartsWith(name + "_", StringComparison.OrdinalIgnoreCase))
                    return name + ".png";
            return string.Equals(type.TypeName, "Вставка гибкая", StringComparison.OrdinalIgnoreCase) ? "Гибкая вставка.png" : null;
        }

        private ImageSource SectionIcon(SectionTypeChoice type)
        {
            string fileName = SectionIconFileName(type);
            if (fileName == null)
                return null;
            ImageSource image;
            if (_sectionIcons.TryGetValue(fileName, out image))
                return image;
            // PNG в Imagens/VentilationConfigurator: Действие при сборке = EmbeddedResource.
            // Префикс пространства имён проекта может отличаться от имени сборки.
            var assembly = typeof(OV_VentConfiguratorForm).Assembly;
            string suffix = "Imagens.VentilationConfigurator." + fileName;
            string resource = assembly.GetManifestResourceNames().FirstOrDefault(name => string.Equals(name, suffix, StringComparison.OrdinalIgnoreCase) || name.EndsWith("." + suffix, StringComparison.OrdinalIgnoreCase));
            image = null;
            if (resource != null)
            {
                try
                {
                    using (Stream stream = assembly.GetManifestResourceStream(resource))
                    {
                        if (stream != null)
                        {
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.DecodePixelWidth = 160;
                            bitmap.StreamSource = stream;
                            bitmap.EndInit();
                            bitmap.Freeze();
                            image = bitmap;
                        }
                    }
                }
                catch (Exception ex) when (ex is IOException || ex is FormatException || ex is NotSupportedException || ex is ArgumentException || ex is InvalidOperationException)
                {
                    // Отсутствующая или повреждённая иконка оставляет пустую область в карточке.
                }
            }

            _sectionIcons[fileName] = image;
            return image;
        }

        internal void RenderBlocks()
        {
            if (Model.Configuration == null)
                return;
            CompositionSummaryTextBlock.Text = "Состав установки";
            CompositionSummaryTextBlock.ToolTip = null;
            if (_installationLengthText != null)
            {
                _installationLengthText.Text = LocalGeometry.Format(Model.Configuration.Geometry.TotalLengthMm) + " мм";
                _installationLengthText.ToolTip = null;
            }

            SectionBlocksPanel.Children.Clear();
            for (int i = 0; i < Model.Configuration.Blocks.Count; i++)
            {
                int index = i;
                int slot = Model.Configuration.SlotAt(i);
                var block = Model.Configuration.Blocks[i];
                bool selected = ReferenceEquals(block, Model.SelectedSection);
                var accent = new SolidColorBrush(Color.FromRgb(52, 121, 197));
                var selectedText = new SolidColorBrush(Color.FromRgb(29, 78, 123));
                var content = new Grid();
                content.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                content.RowDefinitions.Add(new RowDefinition { Height = new GridLength(38) });
                string title = Model.Configuration.BlockTitle(i);
                var label = new StackPanel();
                // Порядок видимых блоков, включая соединители и клапан, не номер параметра семейства.
                label.Children.Add(new Border
                {
                    Width = 22,
                    Height = 22,
                    CornerRadius = new CornerRadius(11),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 5),
                    Background = selected ? accent : new SolidColorBrush(Color.FromRgb(221, 229, 239)),
                    Child = new TextBlock
                    {
                        Text = (index + 1).ToString(),
                        FontSize = 11,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = selected ? Brushes.White : new SolidColorBrush(Color.FromRgb(67, 85, 105)),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    }
                });
                label.Children.Add(new TextBlock
                {
                    Text = title,
                    FontSize = 10,
                    TextAlignment = TextAlignment.Center,
                    Foreground = selected ? selectedText : Brushes.SlateGray,
                    Margin = new Thickness(0, 0, 0, 5)
                });
                label.Children.Add(new TextBlock
                {
                    Text = block.Type == null ? "Состав не задан" : block.Type.TypeName,
                    TextWrapping = TextWrapping.Wrap,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Height = 35,
                    TextAlignment = TextAlignment.Center,
                    Foreground = selected ? selectedText : new SolidColorBrush(Color.FromRgb(32, 51, 74))
                });
                // Одинаковая высота подписи и области иконки сохраняет общий уровень во всех карточках.
                var icon = new Image
                {
                    Source = SectionIcon(block.Type),
                    Width = 56,
                    Height = 56,
                    Stretch = Stretch.Uniform,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false
                };
                RenderOptions.SetBitmapScalingMode(icon, BitmapScalingMode.HighQuality);
                label.Children.Add(new Border { Height = 64, Child = icon });
                var select = new Button
                {
                    Content = label,
                    Style = (Style)FindResource("SectionSelectButtonStyle")
                };
                System.Windows.Automation.AutomationProperties.SetName(select, (index + 1) + ". " + title + ": " + (block.Type?.TypeName ?? "Состав не задан"));
                select.Click += (s, e) =>
                {
                    CurrentOV_VentConfiguratorVM.SelectSection(block);
                };
                if (!block.IsFixed)
                {
                    select.PreviewMouseLeftButtonDown += (s, e) => _dragStart = e.GetPosition(this);
                    select.PreviewMouseMove += (s, e) =>
                    {
                        Point point = e.GetPosition(this);
                        if (e.LeftButton == MouseButtonState.Pressed && (Math.Abs(point.X - _dragStart.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(point.Y - _dragStart.Y) > SystemParameters.MinimumVerticalDragDistance))
                        {
                            DragDrop.DoDragDrop(select, new DataObject("KPLN.Ventilation.Section", block), DragDropEffects.Move);
                            e.Handled = true;
                        }
                    };
                }

                content.Children.Add(select);
                var buttons = new UniformGrid
                {
                    Rows = 2,
                    Columns = 2
                };
                buttons.Children.Add(BlockButton("←", "Переместить влево", !block.IsFixed && index > Model.Configuration.FirstSectionIndex, () => CurrentOV_VentConfiguratorVM.MoveSection(index, index - 1)));
                buttons.Children.Add(BlockButton("→", "Переместить вправо", !block.IsFixed && index < Model.Configuration.Blocks.Count - 2, () => CurrentOV_VentConfiguratorVM.MoveSection(index, index + 1)));
                buttons.Children.Add(BlockButton("+", "Добавить секцию", Model.Configuration.SectionCount < Model.Configuration.Capacity, () => CurrentOV_VentConfiguratorVM.InsertSection(index)));
                buttons.Children.Add(BlockButton("×", "Удалить секцию", !block.IsFixed && Model.Configuration.SectionCount > Model.Configuration.Minimum, () => CurrentOV_VentConfiguratorVM.RemoveSection(index)));
                Grid.SetRow(buttons, 1);
                content.Children.Add(buttons);
                // Сумма рамки и отступа всегда 4: выделение не сдвигает подписи и иконки.
                var border = new Border
                {
                    Child = content,
                    Height = 212,
                    Padding = new Thickness(selected ? 2 : 3),
                    Margin = new Thickness(2, 0, 2, 0),
                    BorderThickness = new Thickness(selected ? 2 : 1),
                    CornerRadius = new CornerRadius(6),
                    BorderBrush = selected ? accent : Brushes.LightSlateGray,
                    Background = selected ? new SolidColorBrush(Color.FromRgb(220, 235, 255)) : block.IsValve ? new SolidColorBrush(Color.FromRgb(237, 247, 247)) : block.IsConnector ? new SolidColorBrush(Color.FromRgb(237, 243, 248)) : Brushes.White,
                    AllowDrop = !block.IsFixed
                };
                border.Drop += (s, e) =>
                {
                    var moved = e.Data.GetData("KPLN.Ventilation.Section") as SectionValue;
                    if (moved == null)
                        return;
                    CurrentOV_VentConfiguratorVM.MoveSection(Model.Configuration.Blocks.IndexOf(moved), index);
                    e.Handled = true;
                };
                var dimensions = new Grid
                {
                    Height = 16,
                    Margin = new Thickness(3, 0, 3, 4)
                };
                var length = DimensionCaption("Длина", block.Length);
                dimensions.Children.Add(length);
                var tooltip = new StackPanel();
                tooltip.Children.Add(new TextBlock { Text = Model.Configuration.Catalog[slot].DisplayName });
                tooltip.Children.Add(new TextBlock { Text = block.Type?.DisplayName, Margin = new Thickness(0, 2, 0, 6) });
                tooltip.Children.Add(DimensionCaption("Ширина", Model.Configuration.BlockWidth(i)));
                tooltip.Children.Add(DimensionCaption("Высота", Model.Configuration.BlockHeight(i)));
                tooltip.Children.Add(DimensionCaption("Длина", block.Length));
                if (block.IsConnector)
                {
                    tooltip.Children.Add(DimensionCaption("Смещение по X", block.OffsetX));
                    tooltip.Children.Add(DimensionCaption("Смещение по Y", block.OffsetY));
                }

                tooltip.Children.Add(new TextBlock { Text = "Размеры в мм", Foreground = Brushes.SlateGray, FontSize = 10 });
                border.ToolTip = tooltip;
                var cell = new StackPanel();
                cell.Children.Add(dimensions);
                cell.Children.Add(border);
                SectionBlocksPanel.Children.Add(cell);
            }
        }

        private static TextBlock DimensionCaption(string label, DimensionValue value)
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.SlateGray));
            var invalid = new DataTrigger
            {
                Binding = new Binding("IsValid")
                {
                    Source = value
                },
                Value = false
            };
            invalid.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.Firebrick));
            style.Triggers.Add(invalid);
            var text = new TextBlock
            {
                FontSize = 10,
                Style = style,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            text.SetBinding(TextBlock.TextProperty, new Binding("Display") { Source = value, StringFormat = label + " - {0}" });
            text.SetBinding(TextBlock.ToolTipProperty, new Binding("Display") { Source = value, StringFormat = label + " - {0} мм" });
            return text;
        }

        private FrameworkElement DimensionEditor(string label, DimensionValue value, string parameterName)
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
            _parameterEditors[FamilyParameterNames.Canonical(parameterName)] = input;
            input.SetBinding(TextBox.TextProperty, ValueBinding(value, "Text"));
            input.SetBinding(TextBox.IsReadOnlyProperty, new Binding("IsReadOnly") { Source = value });
            panel.Children.Add(input);
            return panel;
        }

        private FrameworkElement BooleanEditor(string label, BooleanValue value, string parameterName, bool side)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
            FrameworkElement input;
            if (side)
            {
                var combo = new ComboBox
                {
                    ItemsSource = new[]
                    {
                        "Справа",
                        "Слева"
                    },
                    MinHeight = 34,
                    Style = (Style)FindResource("CleanComboBoxStyle")
                };
                combo.SetBinding(ComboBox.SelectedIndexProperty, ValueBinding(value, "SideIndex"));
                input = combo;
            }
            else
            {
                var check = new CheckBox
                {
                    Content = "Показывать раму",
                    MinHeight = 34,
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                check.SetBinding(CheckBox.IsCheckedProperty, ValueBinding(value, "Value"));
                input = check;
            }

            _parameterEditors[FamilyParameterNames.Canonical(parameterName)] = input;
            input.ToolTip = ParameterTooltip(parameterName, nameOnly: true);
            input.SetBinding(UIElement.IsEnabledProperty, new Binding("IsEnabled") { Source = value });
            panel.Children.Add(input);
            return panel;
        }

        private static FrameworkElement DimensionInfo(string label, DimensionValue value)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = Brushes.SlateGray, Margin = new Thickness(0, 0, 0, 5) });
            var text = new TextBlock
            {
                FontSize = 14,
                Height = 34,
                Padding = new Thickness(0, 8, 0, 0)
            };
            text.SetBinding(TextBlock.TextProperty, new Binding("Display") { Source = value });
            panel.Children.Add(text);
            return panel;
        }

        private ToolTip ParameterTooltip(string name, bool nameOnly = false)
        {
            var dependencies = Model.Configuration?.Catalog.Dependencies;
            string text = nameOnly ? dependencies?.Find(name)?.Name ?? name : dependencies == null ? name : dependencies.Tooltip(name);
            return new ToolTip
            {
                Content = new TextBlock
                {
                    Text = text,
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 740
                }
            };
        }

        private Button BlockButton(string caption, string tooltip, bool enabled, Action action)
        {
            string geometry = caption == "+" ? "M 5 0 L 5 10 M 0 5 L 10 5" : caption == "×" ? "M 1 1 L 9 9 M 9 1 L 1 9" : caption == "←" ? "M 9 5 L 1 5 M 4 2 L 1 5 L 4 8" : "M 1 5 L 9 5 M 6 2 L 9 5 L 6 8";
            var brush = new SolidColorBrush(caption == "+" ? Color.FromRgb(40, 132, 83) : caption == "×" ? Color.FromRgb(195, 63, 69) : Color.FromRgb(83, 101, 122));
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
            button.Click += (s, e) =>
            {
                action();
                e.Handled = true;
            };
            return button;
        }

        internal void RenderSelectedSection()
        {
            SelectedSectionPanel.Children.Clear();
            if (Model.SelectedSection == null)
                return;
            int index = Model.Configuration.Blocks.IndexOf(Model.SelectedSection);
            int slot = Model.Configuration.SlotAt(index);
            SelectedSectionPanel.Children.Add(new TextBlock
            {
                Text = Model.SelectedSection.IsValve ? "Клапан" : Model.SelectedSection.IsConnector ? "Соединитель" : "Секция",
                FontWeight = FontWeights.SemiBold,
                FontSize = 15,
                Margin = new Thickness(0, 0, 0, 10)
            });
            var definition = Model.Configuration.Catalog[slot];
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
            var choices = CurrentOV_VentConfiguratorVM.GetSelectedSectionChoices();
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
            _parameterEditors[InstallationConfiguration.ParameterName(slot)] = selector;
            TextSearch.SetTextPath(selector, "TypeName");
            ScrollViewer.SetHorizontalScrollBarVisibility(selector, ScrollBarVisibility.Disabled);
            selector.SetBinding(ComboBox.SelectedItemProperty, ValueBinding(Model.SelectedSection, "Type"));
            selector.SetBinding(ComboBox.ToolTipProperty, new Binding("Type.DisplayName") { Source = Model.SelectedSection });
            SelectedSectionPanel.Children.Add(selector);
            var sizes = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var length = DimensionEditor("Длина", Model.SelectedSection.Length, InstallationConfiguration.DimensionParameterName(slot, "Длина"));
            length.Width = 150;
            sizes.Children.Add(length);
            if (Model.SelectedSection.IsConnector)
            {
                var width = DimensionEditor("Ширина", Model.SelectedSection.Width, InstallationConfiguration.DimensionParameterName(slot, "Ширина"));
                var height = DimensionEditor("Высота", Model.SelectedSection.Height, InstallationConfiguration.DimensionParameterName(slot, "Высота"));
                var offsetX = DimensionEditor("Смещение по X", Model.SelectedSection.OffsetX, InstallationConfiguration.DimensionParameterName(slot, "Смещение по X"));
                var offsetY = DimensionEditor("Смещение по Y", Model.SelectedSection.OffsetY, InstallationConfiguration.DimensionParameterName(slot, "Смещение по Y"));
                foreach (var field in new[]
                {
                    width,
                    height,
                    offsetX,
                    offsetY
                }

                )
                {
                    field.Width = 150;
                    field.Margin = new Thickness(12, 0, 0, 0);
                }

                sizes.Children.Add(width);
                sizes.Children.Add(height);
                sizes.Children.Add(offsetX);
                sizes.Children.Add(offsetY);
            }
            else
            {
                var width = DimensionInfo("Ширина", Model.Configuration.BlockWidth(index));
                var height = DimensionInfo("Высота", Model.Configuration.BlockHeight(index));
                width.Width = height.Width = 150;
                width.Margin = height.Margin = new Thickness(12, 0, 0, 0);
                sizes.Children.Add(width);
                sizes.Children.Add(height);
            }

            SelectedSectionPanel.Children.Add(sizes);
            _sectionParametersPanel = new WrapPanel
            {
                Margin = new Thickness(0, 14, 0, 0)
            };
            SelectedSectionPanel.Children.Add(_sectionParametersPanel);
            RenderSectionParameters();
        }

        internal void RenderSectionParameters()
        {
            if (_sectionParametersPanel == null)
                return;
            _sectionParametersPanel.Children.Clear();
            if (Model.SelectedSection == null)
                return;
            var names = SharedParameters.ForSection(Model.SelectedSection.Type).ToList();
            _sectionParametersPanel.Visibility = names.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
            foreach (string name in names)
            {
                BooleanValue flag;
                FrameworkElement editor = Model.Configuration.SharedBooleans.TryGetValue(name, out flag) ? BooleanEditor(SharedParameters.Label(name), flag, name, true) : DimensionEditor(SharedParameters.Label(name), Model.Configuration.SharedDimensions[name], name);
                editor.Width = 220;
                editor.Margin = new Thickness(0, 0, 14, 12);
                _sectionParametersPanel.Children.Add(editor);
            }
        }
    }
}
