using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class Furniture2DFrom3D : Window
    {
        private readonly ExternalEvent _selectionEvent;
        private readonly SelectFurnitureElementHandler _selectionHandler;

        public Furniture2DFrom3D(
            ExternalEvent selectionEvent,
            SelectFurnitureElementHandler selectionHandler,
            IEnumerable<FurnitureReplacementResult> results)
        {
            InitializeComponent();
            _selectionEvent = selectionEvent;
            _selectionHandler = selectionHandler;

            Results = new ObservableCollection<FurnitureReplacementResult>(
                results
                    .OrderBy(x => (int)x.Status)
                    .ThenBy(x => x.Name)
                    .ThenBy(x => x.Id));

            DataContext = this;
        }

        public ObservableCollection<FurnitureReplacementResult> Results { get; private set; }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SaveReport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Сохранить отчёт о замене семейств",
                InitialDirectory = Environment.GetFolderPath(
                    Environment.SpecialFolder.DesktopDirectory),
                FileName = "Отчёт_замены_семейств_" +
                           DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt",
                DefaultExt = ".txt",
                Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*",
                AddExtension = true,
                OverwritePrompt = true,
                RestoreDirectory = true
            };

            if (dialog.ShowDialog(this) != true)
                return;

            try
            {
                StringBuilder report = new StringBuilder();
                report.AppendLine("ОТЧЁТ О ЗАМЕНЕ СЕМЕЙСТВ");
                report.AppendLine("Дата: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));
                report.AppendLine();
                report.AppendLine("СТАТУС\tID\tИМЯ\tРАСШИФРОВКА");

                foreach (FurnitureReplacementResult item in
                    ResultGrid.Items.OfType<FurnitureReplacementResult>())
                {
                    report.Append(CleanReportValue(item.StatusText)).Append('\t')
                          .Append(item.Id).Append('\t')
                          .Append(CleanReportValue(item.Name)).Append('\t')
                          .AppendLine(CleanReportValue(item.Error));
                }

                File.WriteAllText(
                    dialog.FileName,
                    report.ToString(),
                    new UTF8Encoding(true));

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
                    "Не удалось сохранить отчёт.\n\n" + ex.Message,
                    "Ошибка сохранения",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static string CleanReportValue(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();
        }

        private void ResultGrid_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            FurnitureReplacementResult selected =
                ResultGrid.SelectedItem as FurnitureReplacementResult;

            if (selected == null)
                return;

            _selectionHandler.SetElementId(selected.Id);
            _selectionEvent.Raise();
        }
    }

    public sealed class SelectFurnitureElementHandler : IExternalEventHandler
    {
        private readonly object _syncRoot = new object();
        private int? _requestedElementId;

        public void SetElementId(int elementId)
        {
            lock (_syncRoot)
                _requestedElementId = elementId;
        }

        public void Execute(UIApplication application)
        {
            int? requestedId;
            lock (_syncRoot)
            {
                requestedId = _requestedElementId;
                _requestedElementId = null;
            }

            if (!requestedId.HasValue || application.ActiveUIDocument == null)
                return;

            Autodesk.Revit.DB.ElementId id =
                new Autodesk.Revit.DB.ElementId(requestedId.Value);

            if (application.ActiveUIDocument.Document.GetElement(id) == null)
                return;

            application.ActiveUIDocument.Selection.SetElementIds(
                new List<Autodesk.Revit.DB.ElementId> { id });
            application.ActiveUIDocument.RefreshActiveView();
        }

        public string GetName()
        {
            return "Выбрать элемент мебели из таблицы результатов";
        }
    }

    public enum FurnitureReplacementStatus
    {
        Failed = 0,
        Warning = 1,
        Success = 2
    }

    public sealed class FurnitureReplacementResult
    {
        public FurnitureReplacementResult(int id, string name)
        {
            Id = id;
            Name = name;
            Status = FurnitureReplacementStatus.Success;
            Error = string.Empty;
        }

        public int Id { get; set; }
        public string Name { get; private set; }
        public FurnitureReplacementStatus Status { get; private set; }
        public string Error { get; private set; }

        public string StatusText
        {
            get
            {
                if (Status == FurnitureReplacementStatus.Failed)
                    return "ОШИБКА";
                if (Status == FurnitureReplacementStatus.Warning)
                    return "ВОЗНИКЛИ ПРОБЛЕМЫ";
                return "OK";
            }
        }

        public void MarkSuccess(string information)
        {
            Status = FurnitureReplacementStatus.Success;
            Error = information ?? string.Empty;
        }

        public void MarkWarning(string warning)
        {
            if (Status != FurnitureReplacementStatus.Failed)
                Status = FurnitureReplacementStatus.Warning;
            AppendMessage(warning);
        }

        public void MarkFailed(string error)
        {
            Status = FurnitureReplacementStatus.Failed;
            AppendMessage(error);
        }

        private void AppendMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            Error = string.IsNullOrWhiteSpace(Error)
                ? message
                : Error + " | " + message;
        }
    }
}