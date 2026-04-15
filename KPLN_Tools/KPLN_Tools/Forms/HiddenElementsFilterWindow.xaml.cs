using Autodesk.Revit.DB;
using KPLN_Tools.ExternalCommands;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class HiddenElementsFilterWindow : Window
    {
        private const string FilterParameterName = "KPLN_Фильтрация";

        private readonly Document _doc;
        private readonly HiddenElementsScanResult _scanResult;

        public HiddenElementsFilterWindow(Document doc, HiddenElementsScanResult scanResult)
        {
            InitializeComponent();

            _doc = doc;
            _scanResult = scanResult ?? new HiddenElementsScanResult();

            SheetsGrid.ItemsSource = _scanResult.Sheets;
            UnhideAllElements = false;
        }

        public bool UnhideAllElements { get; private set; }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            UnhideAllElements = UnhideAllElementsCheckBox.IsChecked == true;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SaveReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Title = "Сохранить отчёт TXT";
                dialog.Filter = "Текстовый файл (*.txt)|*.txt";
                dialog.FileName = "Отчёт_скрытые_элементы_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt";
                dialog.AddExtension = true;
                dialog.DefaultExt = ".txt";

                bool? result = dialog.ShowDialog(this);
                if (result != true)
                    return;

                string reportText = BuildReportText();
                File.WriteAllText(dialog.FileName, reportText, new UTF8Encoding(true));

                MessageBox.Show(
                    this,
                    "Отчёт сохранён:\n" + dialog.FileName,
                    "Сохранение отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "Ошибка при сохранении отчёта:\n" + ex.Message,
                    "Сохранение отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private string BuildReportText()
        {
            StringBuilder sb = new StringBuilder();

            Dictionary<int, List<ViewOccurrenceInfo>> elementViewsMap = BuildElementViewsMap();

            List<PlanHiddenInfo> plansWithHiddenElements = _scanResult.Sheets
                .Where(s => s != null && s.Plans != null)
                .SelectMany(s => s.Plans)
                .Where(p => p != null && p.HiddenElementIds != null && p.HiddenElementIds.Count > 0)
                .OrderBy(p => p.ViewName)
                .ThenBy(p => p.ViewId.IntegerValue)
                .ToList();

            foreach (PlanHiddenInfo plan in plansWithHiddenElements)
            {
                sb.AppendLine((plan.ViewName ?? string.Empty) + " (" + plan.ViewId.IntegerValue + ")");

                List<int> uniqueIdsInPlan = plan.HiddenElementIds
                    .Where(x => x != null)
                    .Select(x => x.IntegerValue)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                foreach (int elementId in uniqueIdsInPlan)
                {
                    List<ViewOccurrenceInfo> otherViews = new List<ViewOccurrenceInfo>();

                    List<ViewOccurrenceInfo> allViewsForElement;
                    if (elementViewsMap.TryGetValue(elementId, out allViewsForElement))
                    {
                        otherViews = allViewsForElement
                            .Where(v => v.ViewId != plan.ViewId.IntegerValue)
                            .OrderBy(v => v.ViewName)
                            .ThenBy(v => v.ViewId)
                            .ToList();
                    }

                    bool hasFilterParameter = ElementHasParameter(elementId, FilterParameterName);

                    StringBuilder lineBuilder = new StringBuilder();
                    lineBuilder.Append("- ");
                    lineBuilder.Append(elementId);

                    if (otherViews.Count > 0)
                    {
                        lineBuilder.Append(" (есть ещё на видах: ");
                        lineBuilder.Append(string.Join("; ", otherViews.Select(v => (v.ViewName ?? string.Empty) + " (" + v.ViewId + ")")));
                        lineBuilder.Append(")");
                    }

                    if (!hasFilterParameter)
                    {
                        lineBuilder.Append(" - нет параметра ");
                        lineBuilder.Append(FilterParameterName);
                    }

                    sb.AppendLine(lineBuilder.ToString());
                }

                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }

        private Dictionary<int, List<ViewOccurrenceInfo>> BuildElementViewsMap()
        {
            Dictionary<int, List<ViewOccurrenceInfo>> result = new Dictionary<int, List<ViewOccurrenceInfo>>();

            List<PlanHiddenInfo> allPlansWithHidden = _scanResult.Sheets
                .Where(s => s != null && s.Plans != null)
                .SelectMany(s => s.Plans)
                .Where(p => p != null && p.HiddenElementIds != null && p.HiddenElementIds.Count > 0)
                .ToList();

            foreach (PlanHiddenInfo plan in allPlansWithHidden)
            {
                List<int> uniqueIdsInPlan = plan.HiddenElementIds
                    .Where(x => x != null)
                    .Select(x => x.IntegerValue)
                    .Distinct()
                    .ToList();

                foreach (int elementId in uniqueIdsInPlan)
                {
                    List<ViewOccurrenceInfo> views;
                    if (!result.TryGetValue(elementId, out views))
                    {
                        views = new List<ViewOccurrenceInfo>();
                        result[elementId] = views;
                    }

                    views.Add(new ViewOccurrenceInfo
                    {
                        ViewId = plan.ViewId.IntegerValue,
                        ViewName = plan.ViewName ?? string.Empty
                    });
                }
            }

            return result;
        }

        private bool ElementHasParameter(int elementId, string parameterName)
        {
            if (_doc == null)
                return false;

            Element element = _doc.GetElement(new ElementId(elementId));
            if (element == null)
                return false;

            Parameter parameter = element.LookupParameter(parameterName);
            return parameter != null;
        }

        private class ViewOccurrenceInfo
        {
            public int ViewId { get; set; }
            public string ViewName { get; set; }
        }
    }
}