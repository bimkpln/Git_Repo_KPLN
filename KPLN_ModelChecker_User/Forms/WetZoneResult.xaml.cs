using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using System;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class WetZoneResult : Window
    {
        private readonly UIDocument _uidoc;
        private readonly string _selectedParam;
        private readonly Dictionary<int, List<Element>> _roomsByFloorParam;

        List<List<Element>> _kitchenOverLiving;
        List<List<Element>> _wetOverLiving;
        List<List<Element>> _kitchenUnderWet;

        public WetZoneResult(UIDocument uidoc, Dictionary<int, List<Element>> roomsByFloorParam,
            List<List<Element>> kitchenOverLiving, List<List<Element>> wetOverLiving, List<List<Element>> kitchenUnderWet, string selectedParam)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _selectedParam = selectedParam;
            _roomsByFloorParam = roomsByFloorParam;

            _kitchenOverLiving = kitchenOverLiving;
            _wetOverLiving = wetOverLiving;
            _kitchenUnderWet = kitchenUnderWet;

            AddViolationBlocks("СП 54.13330.2022 (7.20): Недопустимо размещать мокрые зоны над жилыми помещениями", wetOverLiving);
            AddViolationBlocks("СП 54.13330.2022 (7.20): Недопустимо размещать мокрые зоны над кухнями", kitchenUnderWet);
            AddViolationBlocks("СП 54.13330.2022 (7.21): Недопустимо размещать кухни над жилыми помещениями", kitchenOverLiving);
        }

        // Заполнение интерфейса информацией
        private void AddViolationBlocks(string title, List<List<Element>> rawGroups)
        {
            if (rawGroups.Count == 0)
            {
                return;           
            }
            var grouped = rawGroups
                .GroupBy(g =>
                {
                    var lower = g.First();
                    var uppers = g.Skip(1);
                    var key = string.Join("_", new[] { lower.Id.IntegerValue }
                        .Concat(uppers.Select(u => u.Id.IntegerValue).OrderBy(x => x)));
                    return key;
                })
                .Select(g => g.First())
                .ToList();

            foreach (var group in grouped)
            {
                var lower = group[0];
                var uppers = group.Skip(1).ToList();
                var container = new StackPanel();

                string[] parts = title.Split(new[] { ':' }, 2);
                string mainPart = parts.Length > 0 ? parts[0].Trim() : title;
                string explanation = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                var titleMainText = new TextBlock
                {
                    Text = mainPart,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 0)
                };
                container.Children.Add(titleMainText);
                if (!string.IsNullOrWhiteSpace(explanation))
                {
                    var titleExplanationText = new TextBlock
                    {
                        Text = explanation,
                        FontStyle = FontStyles.Italic,
                        FontSize = 11,
                        Margin = new Thickness(0, 0, 0, 4)
                    };
                    container.Children.Add(titleExplanationText);
                }


                var buttonsPanel = new WrapPanel
                {
                    Orientation = Orientation.Horizontal,  
                    Margin = new Thickness(0, 2, 0, 2)
                };

                int upperFloor = FindFloor(uppers.First());
                foreach (var upper in uppers)
                {
                    string type = GetString(upper, _selectedParam);
                    string kv = GetString(upper, "КВ_Номер");
                    var btn = CreateElementButton(upper, upperFloor, type, kv);
                    buttonsPanel.Children.Add(btn);
                }

                int lowerFloor = FindFloor(lower);
                string lowerType = GetString(lower, _selectedParam);
                string lowerKv = GetString(lower, "КВ_Номер");
                var lowerBtn = CreateElementButton(lower, lowerFloor, lowerType, lowerKv);
                buttonsPanel.Children.Add(lowerBtn);

                container.Children.Add(buttonsPanel);

                // В рамке с фоном
                var border = new Border
                {
                    BorderThickness = new Thickness(1),
                    BorderBrush = System.Windows.Media.Brushes.Gray,
                    Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFFFE0E0")), 
                    Margin = new Thickness(0, 6, 0, 6),
                    Padding = new Thickness(6),
                    CornerRadius = new CornerRadius(6),
                    Child = container
                };

                ViolationsPanel.Children.Add(border);
            }
        }

        // Вспомогательная функция. Поиск номера этажа из общего словаря
        private int FindFloor(Element el)
        {
            foreach (var kvp in _roomsByFloorParam)
            {
                if (kvp.Value.Contains(el))
                    return kvp.Key;
            }
            return -1;
        }

        // Вспомогательная функция. Получение значение параметра (тип помещения)
        private string GetString(Element el, string paramName)
        {
            return el.LookupParameter(paramName)?.AsString() ?? "?";
        }

        // Добавление кнопок
        private Button CreateElementButton(Element el, int florNumber, string type, string kvNum)
        {
            var dock = new DockPanel
            {
                LastChildFill = true
            };
        
            var labelBlock = new TextBlock
            {
                Text = $"{el.Id.IntegerValue} ({type}) ",
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center
            };

            var kvBlock = new TextBlock
            {
                Text = $"[Этаж {florNumber}. Квартира {kvNum}] ",
                FontSize = 11,
                Foreground = Brushes.SteelBlue,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };

            dock.Children.Add(labelBlock);
            dock.Children.Add(kvBlock);
            
            var btn = new Button
            {
                Content = dock,
                Tag = el,
                Height = 23,
                Margin = new Thickness(3),
                Padding = new Thickness(3),
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center
            };

            btn.Click += ElementButton_Click;
            return btn;
        }

        private void ElementButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Element el)
            {
                _uidoc.Selection.SetElementIds(new List<ElementId> { el.Id });
                _uidoc.ShowElements(el.Id);
            }
        }

        // Сохранение отчёта о влажных зонах
        private void TopSaveInfoButton_Click(object sender, RoutedEventArgs e)
        {
            // Строим текст отчёта
            var report = $"Отчёт о проверке мокрых зон по СП 54.13330.2022 документа [{_uidoc.Document.Title}] за {DateTime.Now:dd.MM.yyyy HH:mm}:\n\n";
            report += BuildViolationReport("СП 54.13330.2022 (7.20): Недопустимо размещать мокрые зоны над жилыми помещениями", _wetOverLiving);
            report += BuildViolationReport("СП 54.13330.2022 (7.20): Недопустимо размещать мокрые зоны над кухнями", _kitchenUnderWet);
            report += BuildViolationReport("СП 54.13330.2022 (7.21): Недопустимо размещать кухни над жилыми помещениями", _kitchenOverLiving);

            // Показываем диалог для сохранения файла
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Текстовые файлы (*.txt)|*.txt",
                Title = $"Отчёт по мокрым зонам",
                FileName = $"{_uidoc.Document.Title}{DateTime.Now:ddMMyyyyHHmm}.txt"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    System.IO.File.WriteAllText(saveDialog.FileName, report);
                    MessageBox.Show("Отчёт успешно сохранён.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении файла:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // Вспомогательный метод. Формирование информации
        private string BuildViolationReport(string title, List<List<Element>> groups)
        {
            var result = $"\n{title}:\n";
            if (groups == null || groups.Count == 0)
            {
                result += "В данной категории нарушений не найдено.\n\n";
                return result;
            }

            var uniqueGroups = groups
                .GroupBy(group =>
                {
                    var ids = group.Select(e => e.Id.IntegerValue).OrderBy(x => x);
                    return string.Join("_", ids);
                })
                .Select(g => g.First())
                .ToList();

            int groupIndex = 1;
            foreach (var group in uniqueGroups)
            {
                result += $"  Группа {groupIndex++}:\n";
                var sorted = group
                    .OrderByDescending(e => FindFloor(e))
                    .ToList();

                foreach (var el in sorted)
                {
                    int floor = FindFloor(el);
                    string type = GetString(el, _selectedParam);
                    string kv = GetString(el, "КВ_Номер");
                    result += $"    ID {el.Id.IntegerValue} ({type}) [Этаж {floor}. Квартира {kv}]\n";
                }
            }

            return result;
        }

    }
}