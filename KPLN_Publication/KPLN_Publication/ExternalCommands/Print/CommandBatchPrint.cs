using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker;
using KPLN_Loader.Forms;
using KPLN_Publication.Forms;
using KPLN_Publication.Forms.MVVMCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_Publication.ExternalCommands.Print
{
    [Transaction(TransactionMode.Manual)]
    class CommandBatchPrint : IExternalCommand
    {
        internal const string PluginName = "Пакетная выдача";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Logger logger = new Logger();
                logger.Write("Print started");

                Selection sel = commandData.Application.ActiveUIDocument.Selection;
                Document mainDoc = commandData.Application.ActiveUIDocument.Document;

                string mainDocTitle = EntitySupport.GetDocTitleWithoutRvt(mainDoc.Title);

                //получаю выбранные листы в диспетчере проекта
                View[] selSheets = sel
                    .GetElementIds()
                    .Select(id => mainDoc.GetElement(id))
                    .OfType<ViewSheet>()
                    .Cast<View>()
                    .ToArray();

                // получаю выбранные виды в диспетчере проекта
                View[] selViews = sel
                    .GetElementIds()
                    .Select(id => mainDoc.GetElement(id))
                    .Where(el => el is ViewPlan || el is ViewSection)
                    .Cast<View>()
                    .ToArray();

                if (selSheets.Any() && selViews.Any())
                {
                    message = "Выбраны и листы и виды. Выберите ЛИБО листы, ЛИБО виды в Диспетчере проекта через Shift.";
                    logger.Write("Эксопорт остановлен, выбраны и листы и виды");
                    return Result.Failed;
                }
                else if (!selSheets.Any() && !selViews.Any())
                {
                    message = "Вы не выбрали виды для экспорта. Выберите ЛИБО листы, ЛИБО виды в Диспетчере проекта через Shift.";
                    logger.Write("Эксопорт остановлен, ничего для экспорта не выбрано");
                    return Result.Failed;
                }

                View[] resultViews = selSheets.Any() ? selSheets : selViews;

                //Листы/виды из всех открытых файлов, ключ - имя файла, значение - список видов
                List<string> listErrors = new List<string>();
                FamilyInstance[] tBlocksCollHeap = new FilteredElementCollector(mainDoc)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToArray();
                Dictionary<string, List<MainEntity>> allEntities = EntitySupport.GetAllEntities(commandData, resultViews);
                foreach (View view in resultViews)
                {
                    MainEntity entInBase = allEntities[mainDocTitle].Where(i => i.MainView.Id.Equals(view.Id)).FirstOrDefault();
                    entInBase.IsPrintable = true;

                    if (view is ViewSheet sheet)
                    {
                        // Анализирую листы на наличие 2х основных надписей в одном месте, а также фильтрую пользовательские семейства категории Основная надпись
                        List<Tuple<XYZ, XYZ>> tBlockLocations = new List<Tuple<XYZ, XYZ>>();
                        FamilyInstance[] tBlocksCollHeapOnSheet = tBlocksCollHeap
                            .Where(el => el.OwnerViewId.Equals(view.Id))
                            .ToArray();

                        FamilyInstance[] tBlocksColl = null;
                        if (tBlocksCollHeapOnSheet.Count() >= 1)
                        {
                            tBlocksColl = tBlocksCollHeapOnSheet
                                .Where(fi => fi.Symbol.FamilyName.Contains("Основная надпись") || fi.Symbol.FamilyName.Contains("Основные надписи"))
                                .ToArray();
                        }

                        if (tBlocksColl == null)
                        {
                            listErrors.Add($"{sheet.SheetNumber}-{sheet.Name}: Имеет несколько экземпляров семейства категории \"Основные надписи\", но при этом - нет ни одного экземпляра семейства штампа КПЛН");
                            continue;
                        }

                        // Если штамп 1, то и проверять не нужно. При взятии геометрии BoundingBoxXYZ - лист открывается.
                        if (tBlocksColl.Length <= 1) continue;

                        foreach (FamilyInstance tBlock in tBlocksColl)
                        {
                            if (listErrors.Contains(sheet.Name)) break;

                            List<Tuple<XYZ, XYZ>> templList = new List<Tuple<XYZ, XYZ>>();
                            BoundingBoxXYZ boxXYZ = tBlock.get_BoundingBox(sheet) ?? throw new Exception($"Ошибка в получении BoundingBoxXYZ у листа {sheet.SheetNumber}-{sheet.Name}. Обратись к разработчику!");

                            // Проверяю на дубликаты
                            if (tBlockLocations.Where(tbl => tbl.Item1.IsAlmostEqualTo(boxXYZ.Max, 0.01) && tbl.Item2.IsAlmostEqualTo(boxXYZ.Min, 0.01)).Any())
                                listErrors.Add($"{sheet.SheetNumber}-{sheet.Name}: Несколько экземпляров основной надписи в одном месте, проблема у рамки с Id: {tBlock.Id}");

                            // Добавляю в коллекцию проверенных элементов
                            tBlockLocations.Add(new Tuple<XYZ, XYZ>(boxXYZ.Min, boxXYZ.Max));
                        }

                        // Проверяю на смещение
                        var blocks = tBlockLocations
                            .OrderBy(b => b.Item2.X)
                            .ToList();

                        double tol = 0.001;
                        for (int i = 1; i < blocks.Count; i++)
                        {
                            bool errorXShift = Math.Abs(blocks[i].Item1.X) - Math.Abs(blocks[i - 1].Item2.X) > tol;

                            bool errorY1 = Math.Abs(blocks[i].Item1.Y) - Math.Abs(blocks[i - 1].Item1.Y) > tol;
                            bool errorY2 = Math.Abs(blocks[i].Item2.Y) - Math.Abs(blocks[i - 1].Item2.Y) > tol;

                            bool errorZ1 = Math.Abs(blocks[i].Item1.Z) - Math.Abs(blocks[i - 1].Item1.Z) > tol;
                            bool errorZ2 = Math.Abs(blocks[i].Item2.Z) - Math.Abs(blocks[i - 1].Item2.Z) > tol;

                            if (errorXShift || errorY1 || errorY2 || errorZ1 || errorZ2)
                            {
                                listErrors.Add($"{sheet.SheetNumber}-{sheet.Name}: " +
                                               $"Рамка №{i + 1} не стыкуется с рамкой №{i + 2}. " +
                                               $"Необходимо выровнять без зазоров");
                                break;
                            }
                        }
                    }
                }


                if (listErrors.Count != 0)
                {
                    foreach (string listError in listErrors)
                    {
                        HtmlOutput.Print($"Ошибка: {listError}. Экспорт остановлен", MessageType.Error);
                    }

                    return Result.Failed;
                }

                //очистка старых Schema при необходимости
#if Revit2020 || Debug2020
            try
            {
                Autodesk.Revit.DB.ExtensibleStorage.Schema sch = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(new Guid("414447EA-4228-4B87-A97C-612462722AD4"));
                Autodesk.Revit.DB.ExtensibleStorage.Schema.EraseSchemaAndAllEntities(sch, true);

                Autodesk.Revit.DB.ExtensibleStorage.Schema sch2 = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(new Guid("414447EA-4228-4B87-A97C-612462722AD5"));
                Autodesk.Revit.DB.ExtensibleStorage.Schema.EraseSchemaAndAllEntities(sch2, true);
                logger.Write("Schema очищены");
            }
            catch
            {
                logger.Write("Не удалось очистить Schema");
            }
#endif

                logger.Write($"Пользователь {DBMainService.CurrentDBUser.Id}-{DBMainService.CurrentDBUser.Name}-{DBMainService.CurrentDBUser.Surname}");

                YayPrintSettings printSettings = YayPrintSettings.GetSavedPrintSettings(selSheets.Any());
                FormPrint form = new FormPrint(mainDoc, allEntities, printSettings);
                form.ShowDialog();

                printSettings = form.PrintSettings;
                if (form.DialogResult != System.Windows.Forms.DialogResult.OK || (!printSettings.isPDFExport && !printSettings.isDWGExport))
                    return Result.Cancelled;

                logger.Write("В окне печати нажат ОК, переход к экспорту");

                string msg = string.Empty;
                if (printSettings.isPDFExport && printSettings.isDWGExport)
                {
                    int printedSheetCount = ExportToPDFEXcecute(logger, printSettings, commandData, mainDoc, mainDocTitle, allEntities, form);
                    int dwgExportedSheetCount = ExportToDWGEXcecute(logger, printSettings, commandData, mainDoc, mainDocTitle, allEntities, form);

                    msg = $"Напечатано листов: {printedSheetCount}.\nПереведено в dwg листов: {dwgExportedSheetCount}.";
                    logger.Write($"Экспорт успешно завершен для {printedSheetCount + dwgExportedSheetCount} листа/-ов");
                }
                else if (printSettings.isPDFExport)
                {
                    int printedSheetCount = ExportToPDFEXcecute(logger, printSettings, commandData, mainDoc, mainDocTitle, allEntities, form);

                    msg = $"Напечатано листов: {printedSheetCount}\n";
                    logger.Write($"Напечатано успешно для {printedSheetCount} листа/-ов");
                }
                else if (printSettings.isDWGExport)
                {
                    int dwgExportedSheetCount = ExportToDWGEXcecute(logger, printSettings, commandData, mainDoc, mainDocTitle, allEntities, form);

                    msg = $"Переведено в dwg листов: {dwgExportedSheetCount}";
                    logger.Write($"Экспорт в dwg успешно завершен для {dwgExportedSheetCount} листа/-ов");
                }

                BalloonTip.Show("Экспорт завершен!", msg);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");

                return Result.Failed;
            }
        }

        private int ExportToPDFEXcecute(Logger logger, YayPrintSettings printSettings, ExternalCommandData commandData, Document mainDoc, string mainDocTitle, Dictionary<string, List<MainEntity>> allSheets, FormPrint form)
        {
            string printerName = printSettings.printerName;
            allSheets = form._entitiesSelected;
            logger.Write("Выбранные для печати листы");
            foreach (var kvp in allSheets)
            {
                logger.Write(" Файл " + kvp.Key);
                foreach (MainEntity ms in kvp.Value)
                {
                    logger.Write("  Лист " + ms.MainView.Name);
                }
            }
            string outputFolder = printSettings.outputPDFFolder;

            YayPrintSettings.SaveSettings(printSettings);
            logger.Write("Настройки печати сохранены");

            //Дополнительные возможности работают только с PDFCreator
            if (printerName != "PDFCreator")
            {
                if (printSettings.colorsType == ColorType.MonochromeWithExcludes || printSettings.isMergePdfs || printSettings.isUseOrientation)
                {
                    string errmsg = "Объединение PDF и печать \"Штампа\" в цвете поддерживаются только  для PDFCreator.";
                    errmsg += "\nВо избежание ошибок эти настройки будут отключены.";
                    TaskDialog.Show("Предупреждение", errmsg);
                    printSettings.isMergePdfs = false;
                    printSettings.excludeColors = new List<PdfColor>();
                    printSettings.isUseOrientation = false;
                    logger.Write("Выбранные настройки несовместимы с принтером " + printerName);
                }
            }
            else
            {
                if (!printSettings.isUseOrientation)
                {
                    SupportRegistry.SetOrientationForPdfCreator(OrientationType.Automatic);
                    logger.Write("Установлена ориентация листа Automatic");
                }
            }
            bool printToFile = form.printToFile;
            if (printToFile)
            {
                outputFolder = CreateFolderToEXport(mainDoc, outputFolder, printerName);
                logger.Write("Создана папка для печати: " + outputFolder);
            }

            int printedSheetCount = 0;

            //печатаю листы из каждого выбранного revit-файла
            List<string> pdfFileNames = new List<string>();
            foreach (string docTitle in allSheets.Keys)
            {
                Document openedDoc = null;
                logger.Write("Печать листов из файла " + docTitle);

                RevitLinkType rlt = null;

                //проверяю, текущий это документ или полученный через ссылку
                if (docTitle == mainDocTitle)
                {
                    openedDoc = mainDoc;
                    logger.Write("Это не ссылочный документ");
                }
                else
                {
                    List<RevitLinkType> linkTypes = new FilteredElementCollector(mainDoc)
                        .OfClass(typeof(RevitLinkType))
                        .Cast<RevitLinkType>()
                        .Where(i => EntitySupport.GetDocTitleWithoutRvt(i.Name) == docTitle)
                        .ToList();
                    if (linkTypes.Count == 0) throw new Exception("Cant find opened link file " + docTitle);
                    rlt = linkTypes.First();

                    //проверю, не открыт ли уже документ, который пытаемся печатать
                    foreach (Document testOpenedDoc in commandData.Application.Application.Documents)
                    {
                        if (testOpenedDoc.IsLinked) continue;
                        if (testOpenedDoc.Title == docTitle || testOpenedDoc.Title.StartsWith(docTitle) || docTitle.StartsWith(testOpenedDoc.Title))
                        {
                            openedDoc = testOpenedDoc;
                            logger.Write("Это открытый ссылочный документ");
                        }
                    }

                    //иначе придется открывать документ через ссылку
                    if (openedDoc == null)
                    {
                        logger.Write("Это закрытый ссылочный документ, пытаюсь его открыть");
                        List<Document> linkDocs = new FilteredElementCollector(mainDoc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .Select(i => i.GetLinkDocument())
                            .Where(i => i != null)
                            .Where(i => EntitySupport.GetDocTitleWithoutRvt(i.Title) == docTitle)
                            .ToList();
                        if (linkDocs.Count == 0) throw new Exception("Cant find link file " + docTitle);
                        Document linkDoc = linkDocs.First();

                        if (linkDoc.IsWorkshared)
                        {
                            logger.Write("Это файл совместной работы, открываю с отсоединением");
                            ModelPath mpath = linkDoc.GetWorksharingCentralModelPath();
                            OpenOptions oo = new OpenOptions
                            {
                                DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                            };
                            WorksetConfiguration wc = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                            oo.SetOpenWorksetsConfiguration(wc);
                            rlt.Unload(new SaveCoordinates());
                            openedDoc = commandData.Application.Application.OpenDocumentFile(mpath, oo);
                        }
                        else
                        {
                            logger.Write("Это однопользательский файл");
                            string docPath = linkDoc.PathName;
                            rlt.Unload(new SaveCoordinates());
                            openedDoc = commandData.Application.Application.OpenDocumentFile(docPath);
                        }
                    }
                    logger.Write("Файл-ссылка успешно открыт");
                }


                List<MainEntity> mSheets = allSheets[docTitle];

                if (docTitle != mainDocTitle)
                {
                    List<ViewSheet> linkSheets = new FilteredElementCollector(openedDoc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .ToList();
                    List<MainEntity> tempSheets = new List<MainEntity>();
                    foreach (MainEntity ms in mSheets)
                    {
                        foreach (ViewSheet vs in linkSheets)
                        {
                            if (ms.ViewId.Equals(vs.Id))
                            {
                                MainEntity newMs = new MainEntity(vs);
                                tempSheets.Add(newMs);
                            }
                        }
                    }
                    mSheets = tempSheets;
                }
                logger.Write("Листов для печати найдено в данном файле: " + mSheets.Count.ToString());

                PrintManager pManager = openedDoc.PrintManager;
                pManager.SelectNewPrintDriver(printerName);
                pManager = openedDoc.PrintManager;
                pManager.PrintRange = Autodesk.Revit.DB.PrintRange.Current;
                pManager.Apply();


                //список основных надписей нужен потому, что размеры листа хранятся в них
                //могут быть примечания, сделанные Основной надписью, надо их отфильровать, поэтому >0.6
                List<FamilyInstance> titleBlocks = new FilteredElementCollector(openedDoc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .Cast<FamilyInstance>()
                    .Where(t => t.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() > 0.6)
                    .ToList();
                logger.Write("Найдено основных надписей: " + titleBlocks.Count.ToString());


                //получаю имя формата и проверяю, настроены ли размеры бумаги в Сервере печати
                string formatsCheckinMessage = PrintSupport.PrintFormatsCheckIn(openedDoc, printerName, titleBlocks, ref mSheets, logger);
                if (formatsCheckinMessage != "")
                {
                    logger.Write("Проверка форматов листов неудачна: " + formatsCheckinMessage);
                    HtmlOutput.Print("Проверка форматов листов неудачна: " + formatsCheckinMessage, MessageType.Error);
                    return 0;
                }

                logger.Write("Проверка форматов листов выполнена успешно, переход к печати");


                // Открываю окно статуса
                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel
                {
                    ProcessTitle = "Экспорт в pdf видов/листов...",
                    MaxProgress = mSheets.Count
                };
                ProgressWindow window = null;
                ManualResetEvent formReady = new ManualResetEvent(false);
                Thread uiThread = new Thread(() =>
                {
                    window = new ProgressWindow(progressInfoViewModel);
                    window.Show();
                    formReady.Set();
                    System.Windows.Threading.Dispatcher.Run();
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.IsBackground = true;
                uiThread.Start();
                formReady.WaitOne();


                //печатаю каждый лист
                foreach (MainEntity msheet in mSheets)
                {
                    if (progressInfoViewModel.IsCancellationRequested)
                    {
                        logger.Write("Пользователь отменил процесс экспорта в pdf");
                        break;
                    }

                    logger.Write(" ");
                    logger.Write("Печатается лист: " + msheet.MainView.Name);
                    if (printSettings.isRefreshSchedules)
                    {
                        SchedulesRefresh.Start(openedDoc, msheet.MainView);
                        logger.Write("Спецификации обновлены успешно");
                    }

                    using (Transaction t = new Transaction(openedDoc))
                    {
                        t.Start("Профили печати");

                        string fileName = msheet.NameByConstructor(printSettings.pdfNameConstructor);

                        if (printerName == "PDFCreator" && printSettings.isUseOrientation)
                        {
                            if (msheet.IsVertical)
                            {
                                SupportRegistry.SetOrientationForPdfCreator(OrientationType.Portrait);
                                logger.Write("Принудительно установлена Portrait ориентация");
                            }
                            if (!msheet.IsVertical)
                            {
                                SupportRegistry.SetOrientationForPdfCreator(OrientationType.Landscape);
                                logger.Write("Принудительно установлена Landscape ориентация");
                            }
                        }

                        for (int i = 0; i < msheet.TitleBlocks.Count; i++)
                        {
                            string tempFilename = "";
                            if (msheet.TitleBlocks.Count > 1)
                            {
                                logger.Write("На листе более 1 основной надписи! Печать части №" + i.ToString());
                                tempFilename = fileName.Replace(".pdf", "_" + i.ToString() + ".pdf");
                            }
                            else
                            {
                                logger.Write("На листе 1 основная надпись Id " + msheet.TitleBlocks.First().Id.ToString());
                                tempFilename = fileName;
                            }

                            string fullFilename = System.IO.Path.Combine(outputFolder, tempFilename);
                            fullFilename = fullFilename.Replace("\\\\", "\\");
                            logger.Write("Печать в файл " + fullFilename);
                            pdfFileNames.Add(fullFilename);

                            //смещаю область для печати многолистовых спецификаций
                            double offsetX = -i * msheet.WidthMm / 25.4; //смещение задается в футах!
                            logger.Write("Смещение печати по X: " + offsetX.ToString("F3"));

                            PrintSetting ps = PrintSupport.CreatePrintSetting(openedDoc, pManager, msheet, printSettings, offsetX, 0);

                            pManager.PrintSetup.CurrentPrintSetting = ps;
                            pManager.PrintToFileName = fullFilename;
                            logger.Write("Настройки печати применены, " + ps.Name);

                            pManager.Apply();
                            pManager.SubmitPrint(msheet.MainView);
                            pManager.Apply();

                            logger.Write("Лист успешно отправлен на принтер");
                            msheet.PdfFileName = fullFilename;
                            printedSheetCount++;
                        }

                        if (printerName == "PDFCreator" && printSettings.isUseOrientation)
                            System.Threading.Thread.Sleep(5000);

                        t.RollBack();
                    }

                    ++progressInfoViewModel.CurrentProgress;
                    progressInfoViewModel.DoEvents();
                }

                
                if (!progressInfoViewModel.IsCancellationRequested)
                {
                    window.Dispatcher.Invoke(() =>
                    {
                        progressInfoViewModel.IsComplete = true;
                        window.UpdateLayout();
                    });
                }
    
                if (rlt != null)
                {

                    openedDoc.Close(false);
#if R2017
                    RevitLinkLoadResult LoadResult = rlt.Reload();
#else
                    LinkLoadResult loadResult = rlt.Reload();
#endif
                    logger.Write("Ссылочный документ закрыт");
                }
            }
            //если требуется постобработка файлов - ждем, пока они напечатаются
            if (printSettings.colorsType == ColorType.MonochromeWithExcludes || printSettings.isMergePdfs || printSettings.isExcludeBorders)
            {
                logger.Write(" ");
                logger.Write("Включена постобработка файлов; ожидание окончания печати. Требуемое число файлов " + printedSheetCount);
                int watchTimer = 0;
                while (printToFile)
                {
                    int filescount = System.IO.Directory.GetFiles(outputFolder).Length;
                    logger.Write("Итерация №" + watchTimer + ", файлов напечатано " + filescount);
                    if (filescount == printedSheetCount)
                    {
                        break;
                    }
                    System.Threading.Thread.Sleep(500);
                    watchTimer++;


                    if (watchTimer > 100)
                    {
                        BalloonTip.Show("Обнаружены неполадки", "Печать PDF заняла продолжительное время или произошел сбой. Дождитесь окончания печати.");
                        logger.Write("Не удалось дождаться окончания печати");
                        return 0;
                    }
                }
            }

            List<MainEntity> printedSheets = new List<MainEntity>();
            foreach (List<MainEntity> mss in allSheets.Values)
            {
                printedSheets.AddRange(mss);
            }

            logger.Write("PDF файлы которые должны быть напечатаны:");
            foreach (string pdfname in pdfFileNames)
            {
                logger.Write("  " + pdfname);
            }

            Thread.Sleep(500);
            logger.Write("PDF файлы напечатанные по факту:");
            foreach (string pdfnameOut in System.IO.Directory.GetFiles(outputFolder, "*.pdf"))
            {
                logger.Write("  " + pdfnameOut);
            }

            //преобразую файл с исключением границ при необходимости
            if (printSettings.isExcludeBorders)
            {
                logger.Write("Преобразование PDF файла со скрытием границ");
                foreach (MainEntity msheet in printedSheets)
                {
                    string file = msheet.PdfFileName;
                    string outFile = file.Replace(".pdf", "_OUT_Border.pdf");
                    logger.Write("Файл будет преобразован из " + file + " в " + outFile);

                    PdfWorker.PdfWorker.SetHideColors(printSettings.excludeBorderColors);
                    PdfWorker.PdfWorker.ConvertToBorderHide(file, outFile);

                    System.IO.File.Delete(file);
                    System.IO.File.Move(outFile, file);
                    logger.Write("Лист успешно преобразован со скрытой рамкой цвета '3,2,51'");
                }
            }

            //преобразую файл в черно-белый при необходимости
            if (printSettings.colorsType == ColorType.MonochromeWithExcludes)
            {
                logger.Write("Преобразование PDF файла в черно-белый");
                foreach (MainEntity msheet in printedSheets)
                {
                    if (msheet.ForceColored)
                    {
                        logger.Write("Лист не преобразовывается в черно-белый: " + msheet.MainView.Name);
                        continue;
                    }

                    string file = msheet.PdfFileName;
                    string outFile = file.Replace(".pdf", "_OUT_Color.pdf");
                    logger.Write("Файл будет преобразован из " + file + " в " + outFile);

                    PdfWorker.PdfWorker.SetExcludeColors(printSettings.excludeColors);
                    PdfWorker.PdfWorker.ConvertToGrayScale(file, outFile);

                    //GrayscaleConvertTools.ConvertPdf(file, outFile, ColorType.Grayscale, new List<ExcludeRectangle> { rect, rect2 });

                    File.Delete(file);
                    File.Move(outFile, file);
                    logger.Write("Лист успешно преобразован в ч/б");
                }
            }

            //объединяю файлы при необходимости
            if (printSettings.isMergePdfs)
            {
                logger.Write(" ");
                logger.Write("\nОбъединение PDF файлов");
                System.Threading.Thread.Sleep(500);
                string combinedFile = Path.Combine(outputFolder, mainDoc.Title + ".pdf");

                string[] pdfFileNamesFromPath = Directory.GetFiles(outputFolder, "*.pdf");

                PdfWorker.PdfWorker.CombineMultiplyPDFs(pdfFileNames, combinedFile, logger);

                foreach (string file in pdfFileNames)
                {
                    File.Delete(file);
                    logger.Write("Удален файл " + file);
                }
                logger.Write("Объединено успешно");
            }

            if (printToFile)
            {
                System.Diagnostics.Process.Start(outputFolder);
                logger.Write("Открыта папка " + outputFolder);
            }

            return printedSheetCount;
        }

        private int ExportToDWGEXcecute(Logger logger, YayPrintSettings printSettings, ExternalCommandData commandData, Document mainDoc, string mainDocTitle, Dictionary<string, List<MainEntity>> allSheets, FormPrint form)
        {
            allSheets = form._entitiesSelected;
            logger.Write("Выбранные для экспорта листы\\виды");
            foreach (var kvp in allSheets)
            {
                logger.Write(" Файл " + kvp.Key);
                foreach (MainEntity ms in kvp.Value)
                {
                    logger.Write("  Лист " + ms.MainView.Name);
                }
            }
            string outputFolder = printSettings.outputDWGFolder;

            YayPrintSettings.SaveSettings(printSettings);
            logger.Write("Настройки экспорта сохранены");

            outputFolder = CreateFolderToEXport(mainDoc, outputFolder);
            logger.Write("Создана папка для экспорта: " + outputFolder);

            int exportedSheetCount = 0;

            //печатаю листы из каждого выбранного revit-файла
            foreach (string docTitle in allSheets.Keys)
            {
                Document openedDoc = null;
                logger.Write("Экспорт листов из файла " + docTitle);

                RevitLinkType rlt = null;

                //проверяю, текущий это документ или полученный через ссылку
                if (docTitle == mainDocTitle)
                {
                    openedDoc = mainDoc;
                    logger.Write("Это не ссылочный документ");
                }
                else
                {
                    List<RevitLinkType> linkTypes = new FilteredElementCollector(mainDoc)
                        .OfClass(typeof(RevitLinkType))
                        .Cast<RevitLinkType>()
                        .Where(i => EntitySupport.GetDocTitleWithoutRvt(i.Name) == docTitle)
                        .ToList();
                    if (linkTypes.Count == 0)
                        throw new Exception("Cant find opened link file " + docTitle);

                    rlt = linkTypes.First();

                    //проверю, не открыт ли уже документ, который пытаемся печатать
                    foreach (Document testOpenedDoc in commandData.Application.Application.Documents)
                    {
                        if (testOpenedDoc.IsLinked) continue;
                        if (testOpenedDoc.Title == docTitle || testOpenedDoc.Title.StartsWith(docTitle) || docTitle.StartsWith(testOpenedDoc.Title))
                        {
                            openedDoc = testOpenedDoc;
                            logger.Write("Это открытый ссылочный документ");
                        }
                    }

                    //иначе придется открывать документ через ссылку
                    if (openedDoc == null)
                    {
                        logger.Write("Это закрытый ссылочный документ, пытаюсь его открыть");
                        List<Document> linkDocs = new FilteredElementCollector(mainDoc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .Select(i => i.GetLinkDocument())
                            .Where(i => i != null)
                            .Where(i => EntitySupport.GetDocTitleWithoutRvt(i.Title) == docTitle)
                            .ToList();
                        if (linkDocs.Count == 0) throw new Exception("Cant find link file " + docTitle);
                        Document linkDoc = linkDocs.First();

                        if (linkDoc.IsWorkshared)
                        {
                            logger.Write("Это файл совместной работы, открываю с отсоединением");
                            ModelPath mpath = linkDoc.GetWorksharingCentralModelPath();
                            OpenOptions oo = new OpenOptions
                            {
                                DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                            };
                            WorksetConfiguration wc = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                            oo.SetOpenWorksetsConfiguration(wc);
                            rlt.Unload(new SaveCoordinates());
                            openedDoc = commandData.Application.Application.OpenDocumentFile(mpath, oo);
                        }
                        else
                        {
                            logger.Write("Это однопользательский файл");
                            string docPath = linkDoc.PathName;
                            rlt.Unload(new SaveCoordinates());
                            openedDoc = commandData.Application.Application.OpenDocumentFile(docPath);
                        }
                    }
                    logger.Write("Файл-ссылка успешно открыт");
                }

                List<MainEntity> mSheets = allSheets[docTitle];

                if (docTitle != mainDocTitle)
                {
                    List<ViewSheet> linkSheets = new FilteredElementCollector(openedDoc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .ToList();
                    List<MainEntity> tempSheets = new List<MainEntity>();
                    foreach (MainEntity ms in mSheets)
                    {
                        foreach (ViewSheet vs in linkSheets)
                        {
                            if (ms.ViewId.Equals(vs.Id))
                            {
                                MainEntity newMs = new MainEntity(vs);
                                tempSheets.Add(newMs);
                            }
                        }
                    }
                    mSheets = tempSheets;
                }
                logger.Write("Листов для экспорта найдено в данном файле: " + mSheets.Count.ToString());

                // Создаем набор для экспорта
                ICollection<ElementId> viewsToExport = new List<ElementId>(mSheets.Select(sh => sh.MainView.Id));

                // Открываю окно статуса
                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel
                {
                    ProcessTitle = "Экспорт в dwg видов/листов...",
                    MaxProgress = mSheets.Count
                };
                ProgressWindow window = null;
                ManualResetEvent formReady = new ManualResetEvent(false);
                Thread uiThread = new Thread(() =>
                {
                    window = new ProgressWindow(progressInfoViewModel);
                    window.Show();
                    formReady.Set();
                    System.Windows.Threading.Dispatcher.Run();
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.IsBackground = true;
                uiThread.Start();
                formReady.WaitOne();


                //экспортирую каждый лист
                foreach (MainEntity msheet in mSheets)
                {
                    if (progressInfoViewModel.IsCancellationRequested)
                    {
                        logger.Write("Пользователь отменил процесс экспорта в dwg");
                        break;
                    }

                    logger.Write(" ");
                    logger.Write("Экспортируется лист: " + msheet.MainView.Name);

                    // Настраиваем параметры экспорта
                    DWGExportOptions dwgOptions = form.PrintSettings.dwgExportSettingShell?.DWGExportSetting?.GetDWGExportOptions() ?? new DWGExportOptions();
                    dwgOptions.MergedViews = true;
                    dwgOptions.FileVersion = ACADVersion.R2013;

                    // Экспортируем лист в DWG
                    bool exportSuccess = openedDoc.Export(
                        // Чтобы избежать формирования пути в формате "C:\\DWG_Print\\Test\\Проект1_20.02.2025 11 17 46\-КП_новая-3213719.jpg", нужно текстовый формат пропустить через Path
                        Path.GetFullPath(outputFolder),
                        msheet.NameByConstructor(form.PrintSettings.dwgNameConstructor),
                        new List<ElementId> { msheet.MainView.Id },
                        dwgOptions);
                    if (exportSuccess)
                        exportedSheetCount++;

                    ++progressInfoViewModel.CurrentProgress;
                    progressInfoViewModel.DoEvents();
                }

                if (!progressInfoViewModel.IsCancellationRequested)
                {
                    window.Dispatcher.Invoke(() =>
                    {
                        progressInfoViewModel.IsComplete = true;
                        window.UpdateLayout();
                    });
                }

                if (rlt != null)
                {

                    openedDoc.Close(false);
#if R2017
                    RevitLinkLoadResult LoadResult = rlt.Reload();
#else
                    LinkLoadResult loadResult = rlt.Reload();
#endif
                    logger.Write("Ссылочный документ закрыт");
                }
            }

            System.Diagnostics.Process.Start(outputFolder);
            logger.Write("Открыта папка " + outputFolder);

            return exportedSheetCount;
        }

        private string CreateFolderToEXport(Document doc, string outputFolder, string printerName = null)
        {
            string folder2 = doc.Title + "_" + DateTime.Now.ToString();
            folder2 = folder2.Replace(':', ' ');

            outputFolder = System.IO.Path.Combine(outputFolder, folder2);
            try
            {
                System.IO.Directory.CreateDirectory(outputFolder);
            }
            catch
            {
                return "Невозможно сохранить файлы папку\n" + outputFolder + "\nВыберите другой путь.";
            }

            outputFolder = outputFolder.Replace("\\", "\\\\");

            //пробуем настроить PDFCreator через реестр Windows, для автоматической печати в папку
            if (!string.IsNullOrEmpty(printerName) && printerName == "PDFCreator")
            {
                SupportRegistry.ActivateSettingsForPDFCreator(outputFolder);
            }

            return outputFolder;
        }
    }
}
