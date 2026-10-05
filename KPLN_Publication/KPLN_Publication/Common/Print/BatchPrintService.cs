using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace KPLN_Publication
{
    /// <summary>Публичная пакетная печать без FormPrint, окон прогресса и сообщений.</summary>
    public static class BatchPrintService
    {
        public const int ContractVersion = 1;

        /// <summary>
        /// Вызывать только в контексте Revit API. Все листы текущего проекта, без связей.
        /// Настройки те же, что у FormPrint; неподдержанные режимы отклоняются до печати.
        /// </summary>
        public static object PrintAllSheets(Document doc, YayPrintSettings settings, bool dryRun, int timeoutSeconds)
        {
            if (doc == null || doc.IsFamilyDocument || doc.IsLinked || doc.IsReadOnly || doc.IsModifiable)
                throw new InvalidOperationException("Нужен доступный для временных транзакций документ проекта, без открытой транзакции.");
            if (settings == null || settings.printerName != "PDFCreator" || !settings.isPDFExport || settings.isPrintToPaper || settings.isDWGExport)
                throw new ArgumentException("Без окна поддерживается только PDF через PDFCreator, без печати на бумагу и DWG.");
            if (settings.isMergePdfs || settings.isExcludeBorders || settings.colorsType == ColorType.MonochromeWithExcludes)
                throw new ArgumentException("Объединение, исключение границ и исключения цветов пока не поддерживаются без окна.");
            if (timeoutSeconds < 5 || timeoutSeconds > 60)
                throw new ArgumentOutOfRangeException(nameof(timeoutSeconds), "Ожидание одного PDF: от 5 до 60 секунд.");
            if (!Enum.IsDefined(typeof(ColorType), settings.colorsType) || !Enum.IsDefined(typeof(HiddenLineViewsType), settings.hiddenLineProcessing)
                || !Enum.IsDefined(typeof(RasterQualityType), settings.rasterQuality))
                throw new ArgumentException("Неизвестный режим печати.");

            string root = settings.outputPDFFolder;
            if (string.IsNullOrWhiteSpace(root) || !Path.IsPathRooted(root) || root.StartsWith(@"\\") || root.Contains(".."))
                throw new ArgumentException("Нужен абсолютный локальный путь вывода.");
            root = Path.GetFullPath(root);
            if (!new System.Drawing.Printing.PrinterSettings { PrinterName = settings.printerName }.IsValid)
                throw new InvalidOperationException("Принтер PDFCreator не установлен.");
            PdfCreatorSettingsScope.Validate();

            var sheets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToArray();
            if (sheets.Length == 0 || sheets.Length > 200)
                throw new InvalidOperationException("Для одного запуска требуется от 1 до 200 листов.");

            var entities = sheets.Select(s => new MainEntity(s)).ToList();
            var blocks = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType().Cast<FamilyInstance>()
                .Where(b => (b.get_Parameter(BuiltInParameter.SHEET_HEIGHT)?.AsDouble() ?? 0) > 0.6).ToList();
            var logger = new Logger(_ => { });
            var files = new List<Dictionary<string, object>>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);


            // Проверяем весь комплект до первого задания печати. Ни один лист не пропускается молча.
            foreach (var entity in entities)
            {
                var sheet = (ViewSheet)entity.MainView;
                if (sheet.IsPlaceholder || !sheet.CanBePrinted)
                    throw new InvalidOperationException("Лист нельзя напечатать: " + entity);
                var titleBlocks = blocks.Where(b => b.OwnerViewId.Equals(sheet.Id)).ToArray();
                if (titleBlocks.Length != 1)
                    throw new InvalidOperationException("Для печати без окна нужен один основной штамп на листе: " + entity);
                string check = EntitySupport.CheckTitleblocSizeCorrects(sheet, titleBlocks[0], logger);
                if (!string.IsNullOrEmpty(check))
                    throw new InvalidOperationException(check);
                double width = titleBlocks[0].get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble() * 304.8;
                double height = titleBlocks[0].get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() * 304.8;
                if (PrinterUtility.GetPaperSize(settings.printerName, width, height, logger) == null)
                    throw new InvalidOperationException("В PDFCreator отсутствует формат " + width + "x" + height + " для " + entity);

                string name = entity.NameByConstructor(settings.pdfNameConstructor ?? "");
                if (!name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) || name.Length <= 4 || name.Length > 150
                    || name != Path.GetFileName(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || !names.Add(name))
                    throw new InvalidOperationException("Недопустимое или повторяющееся имя PDF: " + name);
                files.Add(new Dictionary<string, object>
                {
                    { "sheet_id", sheet.Id.ToString() }, { "unique_id", sheet.UniqueId },
                    { "sheet_number", sheet.SheetNumber }, { "sheet_name", sheet.Name }, { "file_name", name },
                    { "status", "planned" }
                });
            }

            var result = new Dictionary<string, object>
            {
                { "contract_version", ContractVersion }, { "dry_run", dryRun }, { "method", "kpln_publication_pdfcreator" },
                { "complete", false }, { "total_sheet_count", sheets.Length }, { "exported_count", 0 },
                { "output_root", root }, { "files", files }, { "settings", new {
                    printer = settings.printerName, name_template = settings.pdfNameConstructor,
                    color = settings.colorsType.ToString(), processing = settings.hiddenLineProcessing.ToString(),
                    raster_quality = settings.rasterQuality.ToString(), refresh_schedules = settings.isRefreshSchedules,
                    use_orientation = settings.isUseOrientation, merge = false, exclude_borders = false, dwg = false } }
            };
            if (dryRun)
                return result;


            // Мьютекс не позволяет двум экземплярам Revit одновременно менять профиль PDFCreator.
            using (var gate = new Mutex(false, @"Local\KPLN_Publication_PDFCreator"))
            {
                bool acquired;
                try { acquired = gate.WaitOne(0); }
                catch (AbandonedMutexException) { acquired = true; }
                if (!acquired)
                    throw new InvalidOperationException("Уже выполняется пакетная печать PDFCreator.");

                try
                {
                    string folder = Path.Combine(root, DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ") + "-" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(folder);
                    result["output_directory"] = folder;
                    result["model_modified_before"] = doc.IsModified;
                    var errors = new List<string>();
                    result["errors"] = errors;
                    var manager = doc.PrintManager;
                    string oldPrinter = manager.PrinterName;
                    PrintRange oldRange = manager.PrintRange;
                    bool oldToFile = manager.PrintToFile;
                    bool? oldCombined = manager.IsVirtual != VirtualPrinterType.None ? (bool?)manager.CombinedFile : null;
                    string oldFile = manager.PrintToFileName;
                    IPrintSetting oldSetting = manager.PrintSetup.CurrentPrintSetting;
                    var restoreInSession = CaptureParameters(manager.PrintSetup.InSession.PrintParameters);

                    using (var group = new TransactionGroup(doc, "KPLN: временная пакетная печать"))
                    {
                        group.Start();
                        try
                        {
                            manager.SelectNewPrintDriver(settings.printerName);
                            ConfigureSingleSheet(manager);
                            manager.Apply();
                            string error = PrintSupport.PrintFormatsCheckIn(doc, settings.printerName, blocks, ref entities, logger, false);
                            if (!string.IsNullOrEmpty(error))
                                throw new InvalidOperationException(error);

                            for (int i = 0; i < entities.Count; i++)
                            {
                                var entity = entities[i];
                                var file = files[i];
                                string path = Path.Combine(folder, (string)file["file_name"]);
                                file["path"] = path;
                                try
                                {
                                    if (settings.isRefreshSchedules)
                                    {
                                        SchedulesRefresh.groupIds.Clear();
                                        try { SchedulesRefresh.Start(doc, entity.MainView, true); }
                                        finally { SchedulesRefresh.groupIds.Clear(); }
                                    }

                                    using (var profile = new PdfCreatorSettingsScope(folder, settings.isUseOrientation, entity.IsVertical))
                                    using (var transaction = new Transaction(doc, "KPLN: временные параметры печати"))
                                    {
                                        transaction.Start();
                                        var options = transaction.GetFailureHandlingOptions();
                                        options.SetFailuresPreprocessor(new SilentPrintFailures());
                                        options.SetClearAfterRollback(true);
                                        transaction.SetFailureHandlingOptions(options);
                                        try
                                        {
                                            var setting = PrintSupport.CreatePrintSetting(doc, manager, entity, settings, 0, 0, false);
                                            manager.PrintSetup.CurrentPrintSetting = setting;
                                            manager.PrintToFileName = path;
                                            manager.Apply();
                                            if (!manager.SubmitPrint(entity.MainView))
                                                throw new IOException("Revit не принял задание печати.");
                                            file["status"] = "submitted";
                                            WaitForPdf(path, timeoutSeconds, file);
                                        }
                                        finally { transaction.RollBack(); }
                                    }
                                    file["status"] = "exported";
                                }
                                catch (Exception ex)
                                {
                                    file["status"] = "failed";
                                    file["error"] = ex.Message;
                                    errors.Add(entity + ": " + ex.Message);
                                    // Не отправляем следующие листы при неизвестном состоянии очереди.
                                    break;
                                }
                            }
                        }
                        catch (Exception ex) { errors.Add(ex.Message); }
                        finally
                        {
                            RestoreStep(() => group.RollBack(), "Откат временных изменений", errors);
                            // Ошибка одного свойства не должна прерывать восстановление остальных.
                            RestoreStep(() => manager.SelectNewPrintDriver(oldPrinter), "Принтер", errors);
                            RestoreStep(() => restoreInSession(manager.PrintSetup.InSession.PrintParameters), "Параметры страницы", errors);
                            RestoreStep(() => RestoreRangeAndCombined(manager, oldRange, oldCombined), "Диапазон печати", errors);
                            RestoreStep(() => RestorePrintToFile(manager, oldToFile, result), "Печать в файл", errors);
                            if (!string.IsNullOrEmpty(oldFile))
                                RestoreStep(() => manager.PrintToFileName = oldFile, "Имя файла", errors);
                            RestoreStep(() => manager.PrintSetup.CurrentPrintSetting = oldSetting, "Профиль печати", errors);
                            RestoreStep(() => manager.Apply(), "Применение настроек", errors);
                        }
                    }

                    int exported = files.Count(f => (string)f["status"] == "exported");
                    result["exported_count"] = exported;
                    result["complete"] = exported == sheets.Length && errors.Count == 0;
                    result["model_modified_after"] = doc.IsModified;
                    return result;
                }
                finally { gate.ReleaseMutex(); }
            }
        }

        private static void ConfigureSingleSheet(PrintManager manager)
        {
            // SubmitPrint(view) печатает один лист за вызов. CombinedFile=true не объединяет разные вызовы.
            // При Current/Visible Revit запрещает явно устанавливать CombinedFile=false.
            if (manager.IsVirtual != VirtualPrinterType.None)
                manager.CombinedFile = true;
            manager.PrintRange = PrintRange.Current;
            manager.PrintToFile = true;
        }

        private static void RestorePrintToFile(PrintManager manager, bool original, Dictionary<string, object> result)
        {
            // Снимок Revit может содержать false даже для PDF24. Повторная установка false
            // запрещена виртуальным драйвером. Не трогаем совпавшее значение, иначе
            // восстанавливаем допустимое состояние и явно сообщаем о нормализации.
            if (manager.PrintToFile == original)
                return;

            if (!original && manager.IsVirtual != VirtualPrinterType.None)
            {
                result["restoration_warnings"] = new[] {
                    "PrintToFile: исходное false недопустимо для виртуального принтера; сохранено true." };
                return;
            }

            manager.PrintToFile = original;
        }

        private static void RestoreRangeAndCombined(PrintManager manager, PrintRange range, bool? combined)
        {
            if (combined.HasValue && manager.IsVirtual != VirtualPrinterType.None)
            {
                // Сначала Select: здесь допустимы оба значения CombinedFile.
                manager.PrintRange = PrintRange.Select;
                manager.CombinedFile = combined.Value;
            }
            manager.PrintRange = range;
        }

        private static void RestoreStep(Action action, string name, List<string> errors)
        {
            try { action(); }
            catch (Exception ex) { errors.Add("Восстановление настроек Revit (" + name + "): " + ex.Message); }
        }

        private static Action<PrintParameters> CaptureParameters(PrintParameters parameters)
        {
            bool crop = parameters.HideCropBoundaries, planes = parameters.HideReforWorkPlanes;
            bool scopes = parameters.HideScopeBoxes, tags = parameters.HideUnreferencedViewTags;
            var zoomType = parameters.ZoomType;
            // В режиме FitToPage Revit запрещает даже чтение процента масштаба.
            int? zoom = zoomType == ZoomType.Zoom ? (int?)parameters.Zoom : null;
            var placement = parameters.PaperPlacement;
            MarginType? margin = placement == PaperPlacementType.Margins ? (MarginType?)parameters.MarginType : null;
            var raster = parameters.RasterQuality;
            var hidden = parameters.HiddenLineViews;
            var color = parameters.ColorDepth;
            var paper = parameters.PaperSize;
            var orientation = parameters.PageOrientation;
            // Отступы доступны только при размещении по пользовательским полям.
            bool hasOffsets = placement == PaperPlacementType.Margins && margin == MarginType.UserDefined;
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            double? x = hasOffsets ? (double?)parameters.UserDefinedMarginX : null;
            double? y = hasOffsets ? (double?)parameters.UserDefinedMarginY : null;
#else
            double? x = hasOffsets ? (double?)parameters.OriginOffsetX : null;
            double? y = hasOffsets ? (double?)parameters.OriginOffsetY : null;
#endif
            return target =>
            {
                target.HideCropBoundaries = crop;
                target.HideReforWorkPlanes = planes;
                target.HideScopeBoxes = scopes;
                target.HideUnreferencedViewTags = tags;
                target.ZoomType = zoomType;
                if (zoom.HasValue)
                    target.Zoom = zoom.Value;
                target.PaperPlacement = placement;
                if (margin.HasValue)
                    target.MarginType = margin.Value;
                target.RasterQuality = raster;
                target.HiddenLineViews = hidden;
                target.ColorDepth = color;
                target.PaperSize = paper;
                target.PageOrientation = orientation;
                if (hasOffsets)
                {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    target.UserDefinedMarginX = x.Value;
                    target.UserDefinedMarginY = y.Value;
#else
                    target.OriginOffsetX = x.Value;
                    target.OriginOffsetY = y.Value;
#endif
                }
            };
        }

        private static void WaitForPdf(string path, int seconds, Dictionary<string, object> result)
        {
            var timer = Stopwatch.StartNew();
            while (timer.Elapsed.TotalSeconds < seconds)
            {
                try
                {
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        byte[] head = new byte[5];
                        if (stream.Read(head, 0, 5) != 5 || Encoding.ASCII.GetString(head) != "%PDF-")
                            throw new IOException("Нет заголовка PDF.");
                        stream.Position = Math.Max(0, stream.Length - 1024);
                        byte[] tail = new byte[stream.Length - stream.Position];
                        stream.Read(tail, 0, tail.Length);
                        if (!Encoding.ASCII.GetString(tail).Contains("%%EOF"))
                            throw new IOException("PDF ещё не завершён.");
                        stream.Position = 0;
                        using (var sha = SHA256.Create())
                            result["sha256"] = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                        result["bytes"] = stream.Length;
                        return;
                    }
                }
                catch (IOException) { Thread.Sleep(250); }
            }
            throw new TimeoutException("PDF не появился за " + seconds + " секунд. Задание может оставаться в очереди; не повторяйте печать автоматически.");
        }

        private sealed class SilentPrintFailures : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                return accessor.GetFailureMessages().Count == 0
                    ? FailureProcessingResult.Continue : FailureProcessingResult.ProceedWithRollBack;
            }
        }
    }
}
