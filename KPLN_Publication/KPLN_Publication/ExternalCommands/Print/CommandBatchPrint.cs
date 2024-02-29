using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_Publication.ExternalCommands.Print
{
    [Transaction(TransactionMode.Manual)]
    class CommandBatchPrint : IExternalCommand
    {
        private static DBUser _currentDBUser;

        internal static DBUser CurrentDBUser
        {
            get
            {
                if (_currentDBUser == null)
                {
                    UserDbService userDbService = (UserDbService)new CreatorUserDbService().CreateService();
                    _currentDBUser = userDbService.GetCurrentDBUser();
                }
                return _currentDBUser;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Module.assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                Logger logger = new Logger();
                logger.Write("Print started");

                Selection sel = commandData.Application.ActiveUIDocument.Selection;
                Document mainDoc = commandData.Application.ActiveUIDocument.Document;

                string mainDocTitle = SheetSupport.GetDocTitleWithoutRvt(mainDoc.Title);

                //листы из всех открытых файлов, ключ - имя файла, значение - список листов
                Dictionary<string, List<MySheet>> allSheets = SheetSupport.GetAllSheets(commandData);

                //получаю выбранные листы в диспетчере проекта
                List<ElementId> selIds = sel.GetElementIds().ToList();
                bool sheetsIsChecked = false;
                List<string> listErrors = new List<string>();
                foreach (ElementId id in selIds)
                {
                    Element elem = mainDoc.GetElement(id);
                    if (!(elem is ViewSheet sheet)) continue;
                    
                    sheetsIsChecked = true;

                    MySheet sheetInBase = allSheets[mainDocTitle].Where(i => i.sheet.Id.IntegerValue == sheet.Id.IntegerValue).First();
                    sheetInBase.IsPrintable = true;

                    // Анализирую листы на наличие 2х основных надписей в одном месте
                    List<Tuple<XYZ, XYZ>> tBlockLocations = new List<Tuple<XYZ, XYZ>>(); 
                    FilteredElementCollector tBlocksColl = new FilteredElementCollector(mainDoc, sheet.Id)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType();

                    foreach(FamilyInstance tBlock in tBlocksColl)
                    {
                        if (listErrors.Contains(sheet.Name)) break;

                        List<Tuple<XYZ, XYZ>> templList = new List<Tuple<XYZ, XYZ>>();
                        BoundingBoxXYZ boxXYZ = tBlock.get_BoundingBox(sheet);
                        if (boxXYZ == null) throw new Exception($"Ошибка в получении BoundingBoxXYZ у листа {sheet.SheetNumber}-{sheet.Name}. Обратись к разработчику!");

                        // Проверяю на дубликаты
                        if (tBlockLocations.Where(tbl => tbl.Item1.IsAlmostEqualTo(boxXYZ.Max, 0.01) && tbl.Item2.IsAlmostEqualTo(boxXYZ.Min, 0.01)).Any())
                            listErrors.Add($"{sheet.SheetNumber}-{sheet.Name}: Несколько экземпляров основной надписи в одном месте, проблема у рамки с Id: {tBlock.Id}");
                        
                        // Добавляю в коллекцию проверенных элементов
                        tBlockLocations.Add(new Tuple<XYZ, XYZ>(boxXYZ.Max, boxXYZ.Min));
                    }

                    // Проверяю на смещение
                    int tBlocksCount = tBlockLocations.Count();
                    if (tBlocksCount > 1)
                    {
                        int cnt = 0;
                        int cntXShift = 0;
                        while (cnt < tBlocksCount)
                        {
                            cntXShift += tBlockLocations.Where(tbl => new XYZ(tbl.Item1.X, tbl.Item2.Y, tbl.Item2.Z).IsAlmostEqualTo(tBlockLocations[cnt].Item2, 0.02)).Count();
                            cnt++;
                        }
                        
                        if (cntXShift != tBlocksCount - 1)
                            listErrors.Add($"{sheet.SheetNumber}-{sheet.Name}: Рамки выставлены со смещением. Необходимо выронять в последовательную цепь без зазоров");
                    }
                }

                if (!sheetsIsChecked)
                {
                    message = "Не выбраны листы. Выберите листы в Диспетчере проекта через Shift.";
                    logger.Write("Печать остановлена, не выбраны листы");
                    return Result.Failed;
                }

                if (listErrors.Count != 0)
                {
                    foreach(string listError in listErrors)
                    {
                        KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput.Print($"Ошибка: {listError}. Печать остановлена", MessageType.Error);
                    }
                    return Result.Failed;
                }


                //запись статистики по файлу
                //ProjectRating.Worker.Execute(commandData);

                //очистка старых Schema при необходимости
#if Revit2020
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

                YayPrintSettings printSettings = YayPrintSettings.GetSavedPrintSettings();
                FormPrint form = new FormPrint(allSheets, printSettings, CurrentDBUser);
                form.ShowDialog();

                if (form.DialogResult != System.Windows.Forms.DialogResult.OK) return Result.Cancelled;
                logger.Write("В окне печати нажат ОК, переход к печати");
                printSettings = form.printSettings;

                string printerName = printSettings.printerName;
                allSheets = form.sheetsSelected;
                logger.Write("Выбранные для печати листы");
                foreach (var kvp in allSheets)
                {
                    logger.Write(" Файл " + kvp.Key);
                    foreach (MySheet ms in kvp.Value)
                    {
                        logger.Write("  Лист " + ms.sheet.Name);
                    }
                }

                string outputFolder = printSettings.outputFolder;

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
                    PrintSupport.CreateFolderToPrint(mainDoc, printerName, ref outputFolder);
                    logger.Write("Создана папка для печати: " + outputFolder);
                }
                //List<string> pfdFileNames = new List<string>();
                int printedSheetCount = 0;

                //печатаю листы из каждого выбранного revit-файла
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
                            .Where(i => SheetSupport.GetDocTitleWithoutRvt(i.Name) == docTitle)
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
                                .Where(i => SheetSupport.GetDocTitleWithoutRvt(i.Title) == docTitle)
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
                                rlt.Unload(new SaveCoordinates());
                                string docPath = linkDoc.PathName;
                                openedDoc = commandData.Application.Application.OpenDocumentFile(docPath);
                            }
                        }
                        logger.Write("Файл-ссылка успешно открыт");
                    }


                    List<MySheet> mSheets = allSheets[docTitle];

                    if (docTitle != mainDocTitle)
                    {
                        List<ViewSheet> linkSheets = new FilteredElementCollector(openedDoc)
                            .OfClass(typeof(ViewSheet))
                            .Cast<ViewSheet>()
                            .ToList();
                        List<MySheet> tempSheets = new List<MySheet>();
                        foreach (MySheet ms in mSheets)
                        {
                            foreach (ViewSheet vs in linkSheets)
                            {
                                if (ms.SheetId == vs.Id.IntegerValue)
                                {
                                    MySheet newMs = new MySheet(vs);
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
                    pManager.PrintRange = PrintRange.Current;
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
                        message = formatsCheckinMessage;
                        logger.Write("Проверка форматов листов неудачна: " + message);
                        return Result.Failed;
                    }
                    logger.Write("Проверка форматов листов выполнена успешно, переход к печати");


                    //печатаю каждый лист
                    foreach (MySheet msheet in mSheets)
                    {
                        logger.Write(" ");
                        logger.Write("Печатается лист: " + msheet.sheet.Name);
                        if (printSettings.isRefreshSchedules)
                        {
                            SchedulesRefresh.Start(openedDoc, msheet.sheet);
                            logger.Write("Спецификации обновлены успешно");
                        }


                        using (Transaction t = new Transaction(openedDoc))
                        {
                            t.Start("Профили печати");

                            string fileName = msheet.NameByConstructor(printSettings.nameConstructor);

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



                            for (int i = 0; i < msheet.titleBlocks.Count; i++)
                            {
                                string tempFilename = "";
                                if (msheet.titleBlocks.Count > 1)
                                {
                                    logger.Write("На листе более 1 основной надписи! Печать части №" + i.ToString());
                                    tempFilename = fileName.Replace(".pdf", "_" + i.ToString() + ".pdf");
                                }
                                else
                                {
                                    logger.Write("На листе 1 основная надпись Id " + msheet.titleBlocks.First().Id.IntegerValue.ToString());
                                    tempFilename = fileName;
                                }

                                string fullFilename = System.IO.Path.Combine(outputFolder, tempFilename);
                                logger.Write("Печать в файл " + fullFilename);

                                //смещаю область для печати многолистовых спецификаций
                                double offsetX = -i * msheet.widthMm / 25.4; //смещение задается в дюймах!
                                logger.Write("Смещение печати по X: " + offsetX.ToString("F3"));

                                PrintSetting ps = PrintSupport.CreatePrintSetting(openedDoc, pManager, msheet, printSettings, offsetX, 0);

                                pManager.PrintSetup.CurrentPrintSetting = ps;
                                logger.Write("Настройки печати применены, " + ps.Name);


                                PrintSupport.PrintView(msheet.sheet, pManager, ps, tempFilename);
                                logger.Write("Лист успешно отправлен на принтер");
                                msheet.PdfFileName = fullFilename;
                                printedSheetCount++;
                            }

                            if (printerName == "PDFCreator" && printSettings.isUseOrientation)
                            {
                                System.Threading.Thread.Sleep(5000);
                            }

                            t.RollBack();


                        }
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
                if (printSettings.colorsType == ColorType.MonochromeWithExcludes || printSettings.isMergePdfs  || printSettings.isExcludeBorders)
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
                            return Result.Failed;
                        }
                    }
                }

                List<MySheet> printedSheets = new List<MySheet>();
                foreach (List<MySheet> mss in allSheets.Values)
                {
                    printedSheets.AddRange(mss);
                }
                List<string> pdfFileNames = printedSheets.Select(i => i.PdfFileName).ToList();
                logger.Write("PDF файлы которые должны быть напечатаны:");
                foreach (string pdfname in pdfFileNames)
                {
                    logger.Write("  " + pdfname);
                }
                logger.Write("PDF файлы напечатанные по факту:");
                foreach (string pdfnameOut in System.IO.Directory.GetFiles(outputFolder, "*.pdf"))
                {
                    logger.Write("  " + pdfnameOut);
                }

                //преобразую файл с исключением границ при необходимости
                if (printSettings.isExcludeBorders)
                {
                    logger.Write("Преобразование PDF файла со скрытием границ");
                    foreach (MySheet msheet in printedSheets)
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
                    foreach (MySheet msheet in printedSheets)
                    {
                        if (msheet.ForceColored)
                        {
                            logger.Write("Лист не преобразовывается в черно-белый: " + msheet.sheet.Name);
                            continue;
                        }

                        string file = msheet.PdfFileName;
                        string outFile = file.Replace(".pdf", "_OUT_Color.pdf");
                        logger.Write("Файл будет преобразован из " + file + " в " + outFile);

                        PdfWorker.PdfWorker.SetExcludeColors(printSettings.excludeColors);
                        PdfWorker.PdfWorker.ConvertToGrayScale(file, outFile);

                        //GrayscaleConvertTools.ConvertPdf(file, outFile, ColorType.Grayscale, new List<ExcludeRectangle> { rect, rect2 });

                        System.IO.File.Delete(file);
                        System.IO.File.Move(outFile, file);
                        logger.Write("Лист успешно преобразован в ч/б");
                    }
                }

                //объединяю файлы при необходимости
                if (printSettings.isMergePdfs)
                {
                    logger.Write(" ");
                    logger.Write("\nОбъединение PDF файлов");
                    System.Threading.Thread.Sleep(500);
                    string combinedFile = System.IO.Path.Combine(outputFolder, mainDoc.Title + ".pdf");

                    PdfWorker.PdfWorker.CombineMultiplyPDFs(pdfFileNames, combinedFile, logger);

                    foreach (string file in pdfFileNames)
                    {
                        System.IO.File.Delete(file);
                        logger.Write("Удален файл " + file);
                    }
                    logger.Write("Объединено успешно");
                }

                if (printToFile)
                {
                    System.Diagnostics.Process.Start(outputFolder);
                    logger.Write("Открыта папка " + outputFolder);
                }

                //восстанавливаю настройки PDFCreator
                //if(printerName == "PDFCreator")
                //{
                //    SupportRegistry.RestoreSettingsForPDFCreator();
                //}


                string msg = "Напечатано листов: " + printedSheetCount;
                BalloonTip.Show("Печать завершена!", msg);
                logger.Write("Печать успешно завершена, напечатано листов " + printedSheetCount);
                return Result.Succeeded;

            }
            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");
                return Result.Failed;
            }
        }
    }
}
