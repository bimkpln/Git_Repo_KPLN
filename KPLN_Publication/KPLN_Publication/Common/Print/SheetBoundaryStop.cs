using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Publication
{
    internal sealed class SheetBoundaryStop
    {
        private readonly Document _document;
        private readonly string _sheetUniqueId;
        private readonly IList<ElementId> _elementIds;
        private readonly ModelPath _reopenPath;
        private readonly bool _detach;

        internal string SheetName { get; }
        internal string Folder { get; set; }
        internal List<string> Files { get; set; } = new List<string>();
        internal int SentCount { get; set; }
        internal bool PrintToFile { get; set; }
        internal bool MergeRequested { get; set; }
        internal string MergedFile { get; set; }
        internal bool Finalized { get; set; }
        internal bool DwgSkipped { get; set; }
        internal string Failure { get; set; }

        internal SheetBoundaryStop(SheetBoundaryIssue issue, ModelPath reopenPath, bool detach)
        {
            _document = issue.Sheet.Document;
            _sheetUniqueId = issue.Sheet.UniqueId;
            _elementIds = issue.ElementIds;
            _reopenPath = reopenPath;
            _detach = detach;
            SheetName = issue.Sheet.SheetNumber + " — " + issue.Sheet.Name;
        }

        internal void OpenSheet(UIApplication application)
        {
            UIDocument uiDocument = application.ActiveUIDocument;
            if (!_document.IsValidObject || uiDocument == null || !uiDocument.Document.Equals(_document))
            {
                if (_document.IsValidObject && !string.IsNullOrEmpty(_document.PathName))
                    uiDocument = application.OpenAndActivateDocument(_document.PathName);
                else if (_reopenPath != null)
                {
                    using (var options = new OpenOptions())
                    {
                        if (_detach)
                            options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                        uiDocument = application.OpenAndActivateDocument(_reopenPath, options, false);
                    }
                }
                else
                    throw new InvalidOperationException("Перейдите в документ и откройте лист вручную: " + SheetName);
            }

            ViewSheet sheet = uiDocument.Document.GetElement(_sheetUniqueId) as ViewSheet
                ?? throw new InvalidOperationException("Не найден лист: " + SheetName);
            uiDocument.ActiveView = sheet;
            uiDocument.Selection.SetElementIds(_elementIds.Where(id => uiDocument.Document.GetElement(id) != null).ToList());
        }

        internal string GetSummary()
        {
            string result = "Печать остановлена перед листом:\n" + SheetName
                + "\n\nЭтот и последующие листы не напечатаны.";
            if (!PrintToFile)
                result += $"\nЗаданий уже отправлено на принтер: {SentCount}.";
            else if (SentCount == 0)
                result += "\nДо остановки не напечатано ни одного файла. Объединять нечего.";
            else
            {
                result += $"\nЗаданий отправлено: {SentCount}.";
                if (!string.IsNullOrEmpty(MergedFile))
                    result += "\nPDF объединены в файл:\n" + MergedFile;
                else if (Finalized)
                    result += $"\nГотовые файлы сохранены отдельно: {Files.Count}.";
                else
                    result += $"\nЗавершённых отдельных PDF сейчас: {Files.Count(PrintedPdfOutput.IsComplete)}."
                        + (MergeRequested ? "\nОбъединение не завершено; исходные файлы сохранены." : "");

                result += "\n\nПосле исправления повторно выберите этот и оставшиеся листы. "
                    + "Новый запуск создаст другую папку. Для полного комплекта объедините PDF из обеих папок вручную.";
            }

            if (PrintToFile && !string.IsNullOrEmpty(Folder))
                result += "\n\nПапка текущей выдачи:\n" + Folder;
            if (DwgSkipped)
                result += "\n\nВыдача DWG не запускалась из-за остановки печати.";
            if (!string.IsNullOrEmpty(Failure))
                result += "\n\nНе удалось завершить обработку:\n" + Failure;
            return result;
        }
    }
}
