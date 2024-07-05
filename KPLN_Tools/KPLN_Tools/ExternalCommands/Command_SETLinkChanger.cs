using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Windows;
using KPLN_Library_Forms.UI.HtmlWindow;
using RevitServerAPILib;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class Command_SETLinkChanger : IExternalCommand
    {
        /// <summary>
        /// Путь к RS#1
        /// </summary>
        private readonly string _rs1Path = "192.168.0.5";
        /// <summary>
        /// Второй путь к RS#1 (сис админ что-то накрутил...)
        /// </summary>
        private readonly string _rs1Path2 = "192.168.20.7";
        /// <summary>
        /// Путь к RS#2
        /// </summary>
        private readonly string _rs2Path = "rs01";
        private readonly int _rsVersion = 2020;
        /// <summary>
        /// Коллекция аббревиатур разделов, которые должны храниться на RS#2
        /// </summary>
        private readonly string[] _fileNameAbrCollForAR = new string[]
        {
            "КР",
            "ОВ",
            "ПТ",
            "ВК",
            "ЭОМ",
            "ПБ",
            "АК",
            "СС",
        };
        /// <summary>
        /// Ссылка на файл со списком моделей с РС№2
        /// </summary>
        private readonly string _pathDBFileRS2 = @"Z:\Отдел BIM\03_Скрипты\СЕТ_Коллеция файлов с РС2.txt";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Подписка на события
            commandData.Application.DialogBoxShowing += OnDialogBoxShowing;

            try
            {
                #region Вывожу окно с запросом на обновление файла-конфига (только для BIM-отдела)
                if (Module.CurrentDBUser.SubDepartmentId == 8)
                {
                    Autodesk.Revit.UI.TaskDialog td = new Autodesk.Revit.UI.TaskDialog("ОШИБКА")
                    {
                        MainIcon = Autodesk.Revit.UI.TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"Обновить файл со списком Revit-файлов с сервера {_rs2Path}?",
                        FooterText = "Занимает много времени, т.к. всё нужно обновить. Лучше запускать этот алгоритм только в случае необходимости (например: в проекте добавились новые файлы)",
                        CommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons.Yes | Autodesk.Revit.UI.TaskDialogCommonButtons.No,
                    };

                    Autodesk.Revit.UI.TaskDialogResult tdResult = td.Show();
                    if (tdResult == Autodesk.Revit.UI.TaskDialogResult.Yes)
                    {
                        // Коллекция моделей РС№2
                        RevitServer rs2 = new RevitServer(_rs2Path, _rsVersion);
                        Task<IList<Model>> rs2ModelsTask = new Task<IList<Model>>(() =>
                        {
                            return GetFilesFromRSFolder(rs2, "Самолет_Сетунь");
                        });
                        rs2ModelsTask.Start();

                        rs2ModelsTask.Wait();

                        // Запись строк в файл
                        using (StreamWriter writer = new StreamWriter(_pathDBFileRS2))
                        {
                            foreach (Model model in rs2ModelsTask.Result)
                            {
                                writer.WriteLine(model.Path);
                            }
                        }
                    }
                }
                #endregion

                Task<List<string>> getDBFilesRS2Task = new Task<List<string>>(() =>
                {
                    List<string> pathes = new List<string>();
                    // Чтение строк из файла
                    using (StreamReader reader = new StreamReader(_pathDBFileRS2))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            pathes.Add(line);
                        }

                        return pathes;
                    }
                });
                getDBFilesRS2Task.Start();

                WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);

                Element[] linkDocColl = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).ToArray();
                int succsesSteps = 0;
                int allModlesForChange = 0;
                foreach (Element linkElem in linkDocColl)
                {
                    if (linkElem is RevitLinkType linkType)
                    {
                        // Анализирую тип связи
                        if (linkType.PathType != PathType.Server)
                            continue;

                        // Анализирую имя проекта на наличие в списке для РС№2 ТОЛЬКО для АР (для остальных - меняются все ссылки)
                        string linkName = linkType.Name;
                        if (doc.Title.Contains("АР") || _fileNameAbrCollForAR.Count(abbr => linkName.Contains(abbr)) == 0)
                            continue;

                        // Анализирую положение связи для РС№2
                        ExternalFileReference extFileRef = linkType.GetExternalFileReference();
                        ModelPath modelPath = extFileRef.GetAbsolutePath();
                        if (!modelPath.CentralServerPath.Contains(_rs1Path) && !modelPath.CentralServerPath.Contains(_rs1Path2))
                            continue;

                        // Генерю новый путь
                        allModlesForChange++;
                        getDBFilesRS2Task.Wait();
                        string currentNewRSModelPath = getDBFilesRS2Task.Result.Where(p => p.Contains(linkName)).FirstOrDefault();
                        if (string.IsNullOrEmpty(currentNewRSModelPath))
                        {
                            HtmlOutput.Print(
                                $"У файла {linkName} не удалось найти дубликат на сервере {_rs2Path}. Скинь в БИМ-отдел, им нужно обновить файл со списком Revit-файлов с сервера {_rs2Path}",
                                MessageType.Error);
                            continue;
                        }

                        ModelPath newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rs2Path}\\{currentNewRSModelPath}");

                        // Проверка РН, что он открыт, иначе - насильно открываю (только через создание элемента в проекте)
                        WorksetId linkWorksetId = linkType.WorksetId;
                        Workset linkWorkset = new FilteredWorksetCollector(doc)
                            .OfKind(WorksetKind.UserWorkset)
                            .Where(ws => ws.Id.IntegerValue == linkWorksetId.IntegerValue)
                            .FirstOrDefault();
                        if (linkWorkset != null
                            && !linkWorkset.IsOpen)
                        {
                            using (Transaction t = new Transaction(doc))
                            {
                                t.Start("KPLN: Открываю РН");//Crating temporary cable tray
                                ElementId typeID = new FilteredElementCollector(doc).OfClass(typeof(Wall)).WhereElementIsElementType().ToElementIds().FirstOrDefault();
                                ElementId levelID = new FilteredElementCollector(doc).OfClass(typeof(Level)).ToElementIds().First();
                                XYZ point_a = new XYZ(-100, 0, 0);
                                XYZ point_b = new XYZ(100, 0, 0); // for start try making a wall in one plane
                                Curve line = Line.CreateBound(point_a, point_b) as Curve;
                                Wall wall = Wall.Create(doc, line, levelID, false);
                                ElementId elementId = wall.Id;

                                //Changing workset of cable tray to workset which we want to open
                                Autodesk.Revit.DB.Parameter wsparam = wall.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                if (wsparam != null && !wsparam.IsReadOnly) wsparam.Set(linkWorkset.Id.IntegerValue);

                                List<ElementId> ids = new List<ElementId>
                                {
                                    elementId
                                };

                                //This command will actualy open workset
                                uidoc.ShowElements(ids);

                                //Delete temporary cable tray
                                doc.Delete(elementId);

                                t.Commit();
                            }
                        }

                        // Обновляю по новому пути
                        linkType.LoadFrom(newModelPath, openConfig);
                        succsesSteps++;
                    }
                }

                HtmlOutput.Print(
                    $"Успешно для {succsesSteps} моделей из {allModlesForChange} моделей, подлежащих замене пути",
                    MessageType.Success);
            }

            catch (System.Exception ex)
            {
                HtmlOutput.Print($"{ex}", MessageType.Error);
            }

            commandData.Application.DialogBoxShowing -= OnDialogBoxShowing;

            return Result.Succeeded;
        }

        /// <summary>
        /// Обработка события всплывающего окна Ревит
        /// </summary>=
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {

            TaskDialogShowingEventArgs taskDialogShowingEventArgs = args as TaskDialogShowingEventArgs;
            if (taskDialogShowingEventArgs.Message.Contains("Не существует открытого вида"))
            {
                args.OverrideResult(5);
            }
            else if (taskDialogShowingEventArgs.Message.Contains("Невозможно подобрать подходящий вид"))
            {
                args.OverrideResult(1);
            }

        }

        private List<Model> GetFilesFromRSFolder(RevitServer rs, string path)
        {
            List<Model> resultModels = new List<Model>();
            IList<Folder> folders = rs.GetFolderContents(path).Folders;
            foreach (Folder folder in folders)
            {
                string newPath = $"{path}\\{folder.Name}";
                IList<Model> models = rs.GetFolderContents(newPath).Models;
                IList<Folder> folderFolders = rs.GetFolderContents(newPath).Folders;
                if (models.Any())
                    resultModels.AddRange(models);

                if (folderFolders.Any())
                    resultModels.AddRange(GetFilesFromRSFolder(rs, folder.Path));

            }

            return resultModels;
        }
    }
}
