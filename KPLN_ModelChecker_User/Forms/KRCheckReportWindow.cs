using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Checks.KR;
using Microsoft.Win32;
using Binding = System.Windows.Data.Binding;
using Panel = System.Windows.Controls.Panel;
using TextBox = System.Windows.Controls.TextBox;
using ComboBox = System.Windows.Controls.ComboBox;

namespace KPLN_ModelChecker_User.Forms
{
    internal sealed class KRCheckReportWindow : Window
    {
        private readonly KRCheckReport _report;
        private readonly DataGrid _table;
        private readonly TextBlock _status;
        private readonly ComboBox _filter;
        private readonly TextBox _search;
        private readonly KRReportSelectionHandler _handler;
        private readonly ExternalEvent _event;

        internal KRCheckReportWindow(UIApplication app, Document doc, string title, KRCheckReport report)
        {
            Title = title;
            Width = 1150; Height = 790; MinWidth = 760; MinHeight = 500;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            new WindowInteropHelper(this).Owner = app.MainWindowHandle;
            _report = report;
            _handler = new KRReportSelectionHandler(doc, message => _status.Text = message);
            _event = ExternalEvent.Create(_handler);
            Closed += (s, e) => { _handler.IsClosed = true; _event.Dispose(); };

            var root = new DockPanel { Margin = new Thickness(14), Background = SystemColors.WindowBrush };
            Content = root;
            var summary = new TextBlock { Text = report.Summary, TextWrapping = TextWrapping.Wrap,
                FontSize = 15, Margin = new Thickness(0, 0, 0, 10) };
            DockPanel.SetDock(summary, Dock.Top); root.Children.Add(summary);
            var filters = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            _filter = new ComboBox { Width = 185, ItemsSource = new[] { "Все замечания", "Ошибка", "Требует разбора", "Не проверено", "Ограничение" }, SelectedIndex = 0 };
            _search = new TextBox { Width = 300, Margin = new Thickness(10, 0, 0, 0), ToolTip = "Поиск по ID, значениям и тексту замечания" };
            filters.Children.Add(_filter); filters.Children.Add(new TextBlock { Text = "Поиск:", Margin = new Thickness(12, 3, 0, 0) }); filters.Children.Add(_search);
            DockPanel.SetDock(filters, Dock.Top); root.Children.Add(filters);

            var bottom = new StackPanel();
            _status = new TextBlock { Text = "Выберите замечание в таблице. Проверки не изменяют модель.", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 6) };
            bottom.Children.Add(_status);
            var buttons = new WrapPanel();
            AddButton(buttons, "Выделить элемент", () => Select("bar", false));
            AddButton(buttons, "Источник параметра", () => Select("source", false));
            AddButton(buttons, "Выделить основу", () => Select("host", false));
            AddButton(buttons, "Все элементы замечания", () => Select("all", false));
            AddButton(buttons, "Показать элемент", () => Select("bar", true));
            AddButton(buttons, "Сохранить отчёт…", SaveReport);
            AddButton(buttons, "Закрыть", Close);
            bottom.Children.Add(buttons); DockPanel.SetDock(bottom, Dock.Bottom); root.Children.Add(bottom);

            _table = new DataGrid { IsReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridSelectionMode.Single,
                SelectionUnit = DataGridSelectionUnit.FullRow, CanUserAddRows = false, EnableRowVirtualization = true };
            Column("Результат", "Kind", 140); Column("ID", "Element", 90);
            Column("Замечание", "Description", 280); Column("Фактически", "Actual", 110);
            Column("Ожидается", "Expected", 110); Column("Основа", "Host", 90); Column("Способ", "Mode", 160);
            root.Children.Add(_table);
            _filter.SelectionChanged += (s, e) => ApplyFilter();
            _search.TextChanged += (s, e) => ApplyFilter();
            ApplyFilter();
        }

        private void Column(string title, string path, double width)
        {
            _table.Columns.Add(new DataGridTextColumn { Header = title, Binding = new Binding(path), Width = width });
        }
        private static void AddButton(Panel panel, string text, Action action)
        {
            var button = new Button { Content = text, Padding = new Thickness(9, 5, 9, 5), Margin = new Thickness(0, 0, 6, 6) };
            button.Click += (s, e) => action(); panel.Children.Add(button);
        }
        private void ApplyFilter()
        {
            string kind = _filter.SelectedItem as string;
            string search = _search.Text.Trim();
            var rows = _report.Rows.Where(r => (kind == "Все замечания" || r.Kind == kind)
                && (search.Length == 0 || (r.Element + " " + r.Details).IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
            _table.ItemsSource = rows;
            _status.Text = "Показано замечаний: " + rows.Count + " из " + _report.Rows.Count + ".";
            if (rows.Count > 0) _table.SelectedIndex = 0;
        }
        private void Select(string target, bool show)
        {
            var row = _table.SelectedItem as KRCheckReportRow;
            if (row == null) { _status.Text = "Выберите замечание."; return; }
            var ids = target == "all" ? row.ElementIds : new[] { target == "bar" ? row.Element : target == "host" ? row.Host : row.SourceId };
            ids = ids.Where(id => !string.IsNullOrEmpty(id)).Distinct().ToArray();
            if (ids.Length == 0) { _status.Text = "Нет экземпляра для выделения: источник может быть типом, а основа — не определена."; return; }
            if (_event.IsPending || _handler.HasRequest) { _status.Text = "Предыдущий запрос выделения ещё выполняется."; return; }
            _handler.SetRequest(ids, show);
            if (_event.Raise() != ExternalEventRequest.Accepted)
            {
                _handler.ClearRequest(); _status.Text = "Revit не принял запрос. Повторите после завершения текущей команды.";
            }
            else _status.Text = "Ожидание выделения в Revit…";
        }
        private void SaveReport()
        {
            var dialog = new SaveFileDialog { Title = "Сохранить полный отчёт", Filter = "Текстовый отчёт (*.txt)|*.txt|Полные данные (*.json)|*.json",
                FileName = "Проверка_КР_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"), AddExtension = true };
            if (dialog.ShowDialog(this) != true) return;
            try
            {
                var content = dialog.FilterIndex == 2 ? _report.Json : _report.Summary + "\r\n\r\n" + _report.Notes + "\r\n\r\n"
                    + string.Join("\r\n--------------------\r\n", _report.Rows.Select(r => r.Kind + "\r\n" + r.Details));
                File.WriteAllText(dialog.FileName, content, new UTF8Encoding(true));
                _status.Text = "Полный отчёт сохранён: " + dialog.FileName;
            }
            catch (Exception ex) { _status.Text = "Не удалось сохранить отчёт: " + ex.Message; }
        }
    }

    internal sealed class KRReportSelectionHandler : IExternalEventHandler
    {
        private readonly Document _document;
        private readonly Action<string> _status;
        private string[] _ids;
        private bool _show;
        internal bool IsClosed;
        internal bool HasRequest => _ids != null;
        internal KRReportSelectionHandler(Document document, Action<string> status) { _document = document; _status = status; }
        internal void SetRequest(string[] ids, bool show) { _ids = ids; _show = show; }
        internal void ClearRequest() { _ids = null; }
        public string GetName() => "Выбор элементов отчёта КР";
        public void Execute(UIApplication app)
        {
            var ids = _ids; _ids = null;
            if (IsClosed || ids == null) return;
            try
            {
                var ui = app.ActiveUIDocument;
                if (!_document.IsValidObject || ui == null || !ui.Document.Equals(_document))
                {
                    _status("Активна другая модель или исходная модель закрыта. Вернитесь в модель отчёта."); return;
                }
                var selected = new List<ElementId>();
                foreach (string value in ids)
                {
                    if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number) || number <= 0) continue;
#if Debug2024 || Revit2024
                    var id = new ElementId(number);
#else
                    if (number > int.MaxValue) continue;
                    var id = new ElementId((int)number);
#endif
                    var element = _document.GetElement(id);
                    if (element != null && !(element is ElementType) && !(element is Material)
                        && !(element is PropertySetElement) && !(element is AppearanceAssetElement)) selected.Add(id);
                }
                if (selected.Count == 0) { _status("Доступные экземпляры не найдены. Возможно, модель изменилась после проверки."); return; }
                ui.Selection.SetElementIds(selected);
                if (_show) ui.ShowElements(selected);
                _status("Выделено элементов: " + selected.Count + ". Пропущено недоступных: " + (ids.Length - selected.Count) + ".");
            }
            catch (Exception ex) { _status("Не удалось выделить / показать элементы: " + ex.Message); }
        }
    }
}
