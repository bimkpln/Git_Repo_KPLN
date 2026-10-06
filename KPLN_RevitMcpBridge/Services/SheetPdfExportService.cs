using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Штатный экспорт листов в PDF для внешнего визуального контроля.</summary>
    internal static class SheetPdfExportService
    {
        internal static object Export(Document doc, IDictionary<string, object> input, string localExportRoot = null)
        {
#if Debug2020 || Revit2020
            throw new BridgeException("unsupported_revit_version", "Штатный PDF-экспорт доступен в Revit 2022+. Для Revit 2020 используйте Пакетную выдачу KPLN вручную.");
#else
            if (doc.IsFamilyDocument || doc.IsLinked)
                throw new BridgeException("not_project_document", "Экспорт требует открытого документа проекта.");

            bool dryRun = input.Flag("dry_run", true);
            ViewSheet[] sheets;
            using (var collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)))
                sheets = collector.Cast<ViewSheet>().OrderBy(s => s.SheetNumber, StringComparer.Ordinal)
                    .ThenBy(s => RevitService.IdValue(s.Id)).ToArray();

            if (sheets.Length == 0)
                throw new BridgeException("no_sheets", "В документе нет листов.");
            if (sheets.Length > 200)
                throw new BridgeException("scope_too_large", "За один вызов допускается до 200 листов; экспорт не запускался.");

            var printable = sheets.Where(s => !s.IsPlaceholder && s.CanBePrinted).ToArray();
            var skipped = sheets.Except(printable).Select(s => new
            {
                sheet = RevitService.ElementInfo(s), sheet_number = s.SheetNumber,
                reason = s.IsPlaceholder ? "placeholder" : "not_printable"
            }).ToArray();

            // Путь выбирает мост: новый уникальный каталог, без перезаписи существующих файлов.
            var exportRoot = localExportRoot ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KPLN", "RevitMcpBridge", "exports");
            var settings = new
            {
                paper_format = "UseSheetSize", scale_percent = 100, color = "Color",
                hide_crop_boundaries = true, hide_scope_boxes = true, hide_reference_planes = true,
                hide_unreferenced_view_tags = false, files = "one_pdf_per_sheet"
            };

            if (dryRun)
                return new
                {
                    dry_run = true, method = "revit_native_pdf_export", output_root = exportRoot, settings,
                    total_sheet_count = sheets.Length, printable_count = printable.Length, skipped,
                    sheets = printable.Select(s => new { sheet = RevitService.ElementInfo(s), sheet_number = s.SheetNumber }).ToArray()
                };

            if (printable.Length == 0)
                throw new BridgeException("no_printable_sheets", "Печатаемых листов нет; файлы не создавались.");

            var folder = Path.Combine(exportRoot, DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            var files = new List<object>();
            var failures = new List<object>();
            bool wasModified = doc.IsModified;


            for (int index = 0; index < printable.Length; index++)
            {
                var sheet = printable[index];
                string fileName = "sheet-" + (index + 1).ToString("D4", CultureInfo.InvariantCulture) + "-" + ModelInspectionService.IdText(sheet.Id);
                string filePath = Path.Combine(folder, fileName + ".pdf");
                try
                {
                    using (var options = new PDFExportOptions())
                    {
                        options.Combine = true;
                        options.FileName = fileName;
                        options.PaperFormat = ExportPaperFormat.Default;
                        options.ColorDepth = ColorDepthType.Color;
                        options.HideCropBoundaries = true;
                        options.HideScopeBoxes = true;
                        options.HideReferencePlane = true;
                        options.HideUnreferencedViewTags = false;
                        options.StopOnError = true;

                        if (!doc.Export(folder, new List<ElementId> { sheet.Id }, options))
                            throw new IOException("Revit сообщил о неуспешном экспорте листа.");
                    }

                    long length;
                    string hash;
                    using (var stream = File.OpenRead(filePath))
                    {
                        length = stream.Length;
                        var signature = new byte[5];
                        if (stream.Read(signature, 0, signature.Length) != 5 || Encoding.ASCII.GetString(signature) != "%PDF-")
                            throw new IOException("Экспортированный файл не содержит заголовок PDF.");

                        stream.Position = 0;
                        using (var sha = SHA256.Create())
                            hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                    }

                    files.Add(new
                    {
                        sheet = RevitService.ElementInfo(sheet), sheet_number = sheet.SheetNumber,
                        path = filePath, bytes = length, sha256 = hash
                    });
                }
                catch (Exception ex) when (ex is Autodesk.Revit.Exceptions.ApplicationException || ex is IOException || ex is UnauthorizedAccessException)
                {
                    failures.Add(new { sheet = RevitService.ElementInfo(sheet), sheet_number = sheet.SheetNumber, path = filePath, message = ex.Message });
                }
            }


            // Манифест связывает каждый PDF с листом, документом и ревизией запроса.
            bool complete = failures.Count == 0 && skipped.Length == 0 && files.Count == sheets.Length;
            var manifestPath = Path.Combine(folder, "manifest.json");
            var result = new
            {
                dry_run = false, method = "revit_native_pdf_export", complete,
                status = complete ? "exported" : "incomplete", output_directory = folder, manifest_path = manifestPath,
                document_id = input.Text("document_id", true), revision_at_request = Json.Integer(input.Get("expected_revision"), "expected_revision"),
                document_title = doc.Title, created_utc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                total_sheet_count = sheets.Length, exported_count = files.Count, skipped, failures, files, settings,
                model_modified_before = wasModified, model_modified_after = doc.IsModified,
                validation = "PDF signature and SHA256 checked; page count, physical dimensions and visual content require external PDF inspection"
            };
            File.WriteAllText(manifestPath, Json.Serialize(result), new UTF8Encoding(false));
            return result;
#endif
        }
    }
}
