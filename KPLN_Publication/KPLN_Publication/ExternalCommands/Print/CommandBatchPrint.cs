using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_DBWorker;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UI.ViewModels;
using KPLN_Loader.Forms;
using KPLN_Publication.Forms;
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

        // Классический MAX_PATH: последний символ зарезервирован под завершающий ноль.
        private const int WindowsMaxPath = 260;
        private const int WindowsMaxDirectoryPath = 248;
        private readonly List<string> _redirectedExportFolders = new List<string>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _redirectedExportFolders.Clear();
            string exportResult = "Экспорт прерван. В указанных папках могут находиться уже выданные файлы.";

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
                    .Where(view => !view.IsTemplate && view.CanBePrinted)
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

                logger.Write($"Пользователь {SQLiteMainService.CurrentDBUser.Id}-{SQLiteMainService.CurrentDBUser.Name}-{SQLiteMainService.CurrentDBUser.Surname}");

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

                exportResult = msg;
                if (_redirectedExportFolders.Count == 0)
                    BalloonTip.Show("Экспорт завершен!", msg);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");

                return Result.Failed;
            }
            finally
            {
                // Показываем фактические папки и при ошибке после частичной выдачи.
                if (_redirectedExportFolders.Count != 0)
                {
                    CustomMessageBox resultWindow = new CustomMessageBox(
                        "Изменена папка экспорта",
                        exportResult + "\n\nПуть превышает допустимую длину Windows. " +
                        "Для экспорта выбраны папки на системном диске:\n\n" +
                        string.Join("\n\n", _redirectedExportFolders) +
                        "\n\nВыделите нужный путь и нажмите Ctrl+C.")
                    {
                        Title = "KPLN: Пакетная выдача",
                        Width = 720,
                        Height = 360,
                        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
                    };
                    WindowHandleSearch.MainWindowHandle.SetAsOwner(resultWindow);
                    resultWindow.ShowDialog();
                }
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
                outputFolder = CreateFolderToEXport(mainDoc, outputFolder, "PDF",
                    GetExportFileNames(mainDoc, allSheets, printSettings, true), logger, printerName);
                logger.Write("Создана папка для печати: " + outputFolder);
            }

            int printedSheetCount = 0;

            //печатаю листы из каждого выбранного revit-файла
            List<string> pdfFileNames = new List<string>();
            // Данные для обработки PDF сохраняем до закрытия связанных документов.
            List<Tuple<string, bool, bool, string>> printedFiles = new List<Tuple<string, bool, bool, string>>();
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
                    mSheets = mSheets.Select(entity =>
                    {
                        View view = openedDoc.GetElement(entity.ViewId) as View;
                        if (view == null)
                            throw new InvalidOperationException("Не найден лист/вид в открытой связи: " + entity.ViewId);
                        return new MainEntity(view);
                    }).ToList();
                }
                logger.Write("Листов для печати найдено в данном файле: " + mSheets.Count.ToString());

                PrintManager pManager = openedDoc.PrintManager;
                pManager.SelectNewPrintDriver(printerName);
                pManager = openedDoc.PrintManager;
                pManager.PrintRange = Autodesk.Revit.DB.PrintRange.Current;
                pManager.PrintToFile = printToFile;
                pManager.Apply();


                //список основных надписей нужен потому, что размеры листа хранятся в них
                //могут быть примечания, сделанные Основной надписью, надо их отфильровать, поэтому >0.6
                List<FamilyInstance> titleBlocks = mSheets.Any(entity => entity.MainView is ViewSheet)
                    ? new FilteredElementCollector(openedDoc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .Cast<FamilyInstance>()
                    .Where(t => t.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() > 0.6)
                    .ToList() : new List<FamilyInstance>();
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
                ProgressWindow window = ProgressWindow.ShowInBackground(progressInfoViewModel, "KPLN: Пакетная выдача");


                try
                {
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
                        if (!msheet.MainView.CanBePrinted)
                            throw new InvalidOperationException("Вид недоступен для печати: " + msheet.MainView.Name);
                        if (msheet.MainView is ViewSheet && printSettings.isRefreshSchedules)
                        {
                            SchedulesRefresh.Start(openedDoc, msheet.MainView);
                            logger.Write("Спецификации обновлены успешно");
                        }

                        using (Transaction t = new Transaction(openedDoc))
                        {
                            t.Start("Профили печати");

                            string fileName = FileNameTemplate.WithExtension(msheet.NameByConstructor(printSettings.pdfNameConstructor), ".pdf");

                            if (msheet.MainView is ViewSheet && printerName == "PDFCreator" && printSettings.isUseOrientation)
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

                            bool isSheet = msheet.MainView is ViewSheet;
                            int partCount = isSheet ? msheet.TitleBlocks.Count : 1;
                            for (int i = 0; i < partCount; i++)
                            {
                                string tempFilename = "";
                                if (partCount > 1)
                                {
                                    logger.Write("На листе более 1 основной надписи! Печать части №" + i.ToString());
                                    tempFilename = fileName.Replace(".pdf", "_" + i.ToString() + ".pdf");
                                }
                                else
                                {
                                    if (isSheet)
                                        logger.Write("На листе 1 основная надпись Id " + msheet.TitleBlocks.First().Id.ToString());
                                    tempFilename = fileName;
                                }

                                string fullFilename = System.IO.Path.Combine(outputFolder, tempFilename);
                                if (printToFile)
                                    CheckExportFilePath(fullFilename);
                                logger.Write("Печать в файл " + fullFilename);
                                pdfFileNames.Add(fullFilename);

                                //смещаю область для печати многолистовых спецификаций
                                double offsetX = isSheet ? -i * msheet.WidthMm / 25.4 : 0; //смещение задается в футах!
                                logger.Write("Смещение печати по X: " + offsetX.ToString("F3"));

                                PrintSetting ps = PrintSupport.CreatePrintSetting(openedDoc, pManager, msheet, printSettings, offsetX, 0);

                                pManager.PrintSetup.CurrentPrintSetting = ps;
                                pManager.PrintToFileName = fullFilename;
                                logger.Write("Настройки печати применены, " + ps.Name);

                                pManager.Apply();
                                if (!pManager.SubmitPrint(msheet.MainView))
                                    throw new InvalidOperationException("Не удалось напечатать лист/вид: " + msheet.MainView.Name);
                                pManager.Apply();

                                logger.Write("Лист успешно отправлен на принтер");
                                msheet.PdfFileName = fullFilename;
                                printedFiles.Add(Tuple.Create(fullFilename, isSheet, msheet.ForceColored, msheet.MainView.Name));
                                printedSheetCount++;
                            }

                            if (msheet.MainView is ViewSheet && printerName == "PDFCreator" && printSettings.isUseOrientation)
                                System.Threading.Thread.Sleep(5000);

                            t.RollBack();
                        }

                        ++progressInfoViewModel.CurrentProgress;
                    }


                    if (!progressInfoViewModel.IsCancellationRequested)
                        progressInfoViewModel.IsComplete = true;
                }
                catch
                {
                    window.CloseSafely();
                    throw;
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
            if (!printToFile)
                return printedSheetCount;

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
                foreach (var printedFile in printedFiles.Where(file => file.Item2))
                {
                    string file = printedFile.Item1;
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
                foreach (var printedFile in printedFiles)
                {
                    if (printedFile.Item3)
                    {
                        logger.Write("Лист/вид не преобразовывается в черно-белый: " + printedFile.Item4);
                        continue;
                    }

                    string file = printedFile.Item1;
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
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputFolder) { UseShellExecute = true });
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

            outputFolder = CreateFolderToEXport(mainDoc, outputFolder, "DWG",
                GetExportFileNames(mainDoc, allSheets, printSettings, false), logger);
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
                    mSheets = mSheets.Select(entity =>
                    {
                        View view = openedDoc.GetElement(entity.ViewId) as View;
                        if (view == null)
                            throw new InvalidOperationException("Не найден лист/вид в открытой связи: " + entity.ViewId);
                        return new MainEntity(view);
                    }).ToList();
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
                ProgressWindow window = ProgressWindow.ShowInBackground(progressInfoViewModel, "KPLN: Пакетная выдача");


                try
                {
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
                        string fileName = FileNameTemplate.WithoutExtension(msheet.NameByConstructor(form.PrintSettings.dwgNameConstructor), ".dwg");
                        CheckExportFilePath(Path.Combine(outputFolder, GetDwgFileName(fileName)));
                        bool exportSuccess = openedDoc.Export(
                            // Чтобы избежать формирования пути в формате "C:\\DWG_Print\\Test\\Проект1_20.02.2025 11 17 46\-КП_новая-3213719.jpg", нужно текстовый формат пропустить через Path
                            Path.GetFullPath(outputFolder),
                            fileName,
                            new List<ElementId> { msheet.MainView.Id },
                            dwgOptions);
                        if (exportSuccess)
                            exportedSheetCount++;

                        ++progressInfoViewModel.CurrentProgress;
                    }

                    if (!progressInfoViewModel.IsCancellationRequested)
                        progressInfoViewModel.IsComplete = true;
                }
                catch
                {
                    window.CloseSafely();
                    throw;
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

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputFolder) { UseShellExecute = true });
            logger.Write("Открыта папка " + outputFolder);

            return exportedSheetCount;
        }

        private static string GetDwgFileName(string fileName)
        {
            return FileNameTemplate.WithExtension(fileName, ".dwg");
        }

        private static List<string> GetExportFileNames(Document doc,
            Dictionary<string, List<MainEntity>> allSheets, YayPrintSettings settings, bool isPdf)
        {
            List<string> fileNames = new List<string>();
            foreach (List<MainEntity> sheets in allSheets.Values)
            {
                if (sheets.Count == 0)
                    continue;

                // Те же основные надписи, которые используются при подготовке печати.
                ILookup<string, FamilyInstance> titleBlocks = isPdf && sheets.Any(entity => entity.MainView is ViewSheet)
                    ? new FilteredElementCollector(sheets[0].MainView.Document)
                        .WhereElementIsNotElementType()
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .Cast<FamilyInstance>()
                        .Where(t => t.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() > 0.6)
                        .ToLookup(t => t.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString())
                    : null;

                foreach (MainEntity sheet in sheets)
                {
                    string fileName = sheet.NameByConstructor(isPdf
                        ? settings.pdfNameConstructor
                        : settings.dwgNameConstructor);
                    fileName = FileNameTemplate.WithExtension(fileName, isPdf ? ".pdf" : ".dwg");
                    if (!isPdf)
                    {
                        fileNames.Add(GetDwgFileName(fileName));
                        continue;
                    }

                    int partCount = sheet.MainView is ViewSheet viewSheet
                        ? titleBlocks[viewSheet.SheetNumber].Count()
                        : 0;
                    for (int i = 0; i < Math.Max(1, partCount); i++)
                    {
                        string partName = partCount > 1
                            ? fileName.Replace(".pdf", "_" + i + ".pdf")
                            : fileName;
                        fileNames.Add(partName);
                        if (sheet.MainView is ViewSheet && settings.isExcludeBorders)
                            fileNames.Add(partName.Replace(".pdf", "_OUT_Border.pdf"));
                        if (settings.colorsType == ColorType.MonochromeWithExcludes)
                            fileNames.Add(partName.Replace(".pdf", "_OUT_Color.pdf"));
                    }
                }
            }

            if (isPdf && settings.isMergePdfs)
                fileNames.Add(doc.Title + ".pdf");

            return fileNames;
        }

        private static bool IsExportPathTooLong(string folder, IEnumerable<string> fileNames)
        {
            return folder.Length >= WindowsMaxDirectoryPath
                || fileNames.Any(name => Path.Combine(folder, name).Length >= WindowsMaxPath);
        }

        private static void CheckExportFilePath(string filePath)
        {
            if (filePath.Length >= WindowsMaxPath || Path.GetFullPath(filePath).Length >= WindowsMaxPath)
                throw new PathTooLongException("Полный путь должен быть короче 260 символов. " +
                    "Сократите имя файла в конструкторе имён:\n" + filePath);
        }

        private string CreateFolderToEXport(Document doc, string outputFolder, string exportType,
            List<string> fileNames, Logger logger, string printerName = null)
        {
            string folderName = doc.Title + "_" + DateTime.Now.ToString("dd.MM.yyyy HH mm ss");
            string requestedFolder = Path.Combine(outputFolder, folderName);
            bool redirect;
            try
            {
                outputFolder = Path.GetFullPath(requestedFolder);
                redirect = IsExportPathTooLong(outputFolder, fileNames);
            }
            catch (PathTooLongException)
            {
                redirect = true;
            }

            if (redirect)
            {
                string systemRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System));
                // Короткая отдельная папка для запуска, без длинного названия модели.
                string runFolder = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                outputFolder = Path.Combine(systemRoot, exportType, runFolder);
                logger.Write("Превышена допустимая длина пути: " + requestedFolder);
                logger.Write("Новая папка " + $"{exportType}_Print" + ": " + outputFolder);
            }

            // Даже короткий корень не поможет, если само имя файла слишком длинное.
            foreach (string fileName in fileNames)
                CheckExportFilePath(Path.Combine(outputFolder, fileName));

            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                throw new IOException("Не удалось создать папку экспорта:\n" + outputFolder +
                    "\nПроверьте доступность диска и права записи.", ex);
            }

            if (redirect)
                _redirectedExportFolders.Add(exportType + ":\n" + outputFolder);

            //пробуем настроить PDFCreator через реестр Windows, для автоматической печати в папку
            if (!string.IsNullOrEmpty(printerName) && printerName == "PDFCreator")
            {
                SupportRegistry.ActivateSettingsForPDFCreator(outputFolder);
            }

            return outputFolder;
        }
    }
}
