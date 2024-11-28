using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class Command_SETLinkChanger : IExternalCommand
    {
        private static DBRevitDialog[] _dbRevitDialogs = null;

        private readonly string _serverOLDPath = "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\";
        private readonly string _rsOLDPath = "192.168.0.5";
        private readonly string _rsOLDRKPath = "rs01";

        private readonly string _rsARPath = "rs03";
        private readonly string _rsKRPath = "rs04";
        private readonly string _rsIOSPath = "rs05";


        internal protected static UIControlledApplication RevitUIControlledApp { get; set; }

        internal static void SetStaticEnvironment(UIControlledApplication application)
        {
            RevitUIControlledApp = application;
        }


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_dbRevitDialogs == null)
            {
                RevitDialogDbService currentRevitDialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
                _dbRevitDialogs = currentRevitDialogDbService.GetDBRevitDialogs().ToArray();
            }

            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Подписка на события
            commandData.Application.DialogBoxShowing += OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.FailuresProcessing += OnFailureProcessing;

            try
            {
                string newModelPathString = string.Empty;
                WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);

                Element[] linkDocColl = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).ToArray();
                int succsesSteps = 0;
                int allModlesForChange = 0;
                foreach (Element linkElem in linkDocColl)
                {

                    if (linkElem is RevitLinkType linkType)
                    {
                        //// Ручное регулирование обрабатываемых связей (перенос будет по частям - поэтому хардкодим)
                        //if (!linkElem.Name.Contains("РФ") && !linkElem.Name.Contains("АР") && !linkElem.Name.Contains("КР"))
                        //    continue;

                        ModelPath newModelPath = null;
                        newModelPathString = string.Empty;
                        try
                        {
                            // Анализирую положение связи 
                            ExternalFileReference extFileRef = linkType.GetExternalFileReference();
                            ModelPath oldModelPath = extFileRef.GetAbsolutePath();
                            string oldModelPathString = ModelPathUtils.ConvertModelPathToUserVisiblePath(oldModelPath);

                            if (!oldModelPathString.Contains(_rsOLDPath) && !oldModelPathString.Contains(_rsOLDRKPath) && !oldModelPathString.Contains(_serverOLDPath))
                                continue;

                            allModlesForChange++;

                            #region Обрабатываю пути АР и файлов ОВ, которые остались на старом серваке
                            if (oldModelPathString.Contains(_rsOLDPath) && (oldModelPathString.Contains("_АР_") || oldModelPathString.Contains("_РФ")))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsARPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath) && oldModelPathString.Contains("_ОВ"))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }
                            #endregion

                            #region Обрабатываю пути КР
                            else if (oldModelPathString.Contains(_rsOLDRKPath))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDRKPath, $@"{_rsKRPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }
                            #endregion

                            #region Обрабатываю пути инженерки
                            // ЭОМ
                            else if (oldModelPathString.Contains(_serverOLDPath) && oldModelPathString.Contains("_ЭОМ"))
                            {
                                if (oldModelPathString.Contains("К1"))
                                    newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.1.ЭОМ\\1.RVT\\К1\\");
                                else if (oldModelPathString.Contains("К2"))
                                    newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.1.ЭОМ\\1.RVT\\К2\\");
                                else if (oldModelPathString.Contains("К3"))
                                    newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.1.ЭОМ\\1.RVT\\К3\\");
                                else if (oldModelPathString.Contains("СТЛ"))
                                    newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.1.ЭОМ\\1.RVT\\СТЛ\\");

                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rsIOSPath}\\Самолет_Сетунь\\ЭОМ\\{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath)
                                && oldModelPathString.Contains("_ЭОМ"))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }

                            //  ВК
                            else if (oldModelPathString.Contains(_serverOLDPath)
                                && oldModelPathString.Contains("_ВК"))
                            {
                                newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.2.ВК\\1.RVT\\");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rsIOSPath}\\Самолет_Сетунь\\ВК\\{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath)
                                && oldModelPathString.Contains("_ВК")
                                && !oldModelPathString.Contains("_КР"))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }

                            // АУПТ
                            else if (oldModelPathString.Contains(_serverOLDPath)
                                && oldModelPathString.Contains("_ПТ"))
                            {
                                newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.3.АУПТ\\1.RVT\\");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rsIOSPath}\\Самолет_Сетунь\\АУПТ\\{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath)
                                && oldModelPathString.Contains("_ПТ"))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }

                            // ОВ
                            else if (oldModelPathString.Contains(_serverOLDPath)
                                && oldModelPathString.Contains("_ОВ"))
                            {
                                newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.4.ОВ\\1.RVT\\");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rsIOSPath}\\Самолет_Сетунь\\ОВ\\{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath)
                                && oldModelPathString.Contains("_ОВ"))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }

                            // СС
                            else if (oldModelPathString.Contains(_serverOLDPath)
                                && (oldModelPathString.Contains("_СС") || oldModelPathString.Contains("_ПБ") || oldModelPathString.Contains("_АК")))
                            {
                                newModelPathString = RemoveSubstringIfExists(oldModelPathString, $"{_serverOLDPath}7.5.СС\\1.RVT\\");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"RSN:\\\\{_rsIOSPath}\\Самолет_Сетунь\\СС\\{newModelPathString}");
                            }
                            else if (oldModelPathString.Contains(_rsOLDPath)
                                && (oldModelPathString.Contains("_СС") || oldModelPathString.Contains("_ПБ") || oldModelPathString.Contains("_АК")))
                            {
                                newModelPathString = oldModelPathString.Replace(_rsOLDPath, $@"{_rsIOSPath}");
                                newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath($"{newModelPathString}");
                            }
                            #endregion

                            if (newModelPath == null)
                            {
                                HtmlOutput.Print(
                                    $"Для пути {oldModelPathString} не подобралось нужного алгоритма замены. Либо файл уже заменен на актуальный (актуальные РС: АР - {_rsARPath}, КР - {_rsKRPath}, ИОС - {_rsIOSPath}), либо нужно скинуть ошибку разработчику",
                                    MessageType.Error);
                                continue;
                            }

                            if (!newModelPath.IsValidObject)
                            {
                                HtmlOutput.Print(
                                    $"Для пути {oldModelPathString} переопределился новый путь {newModelPathString}, но это не путь к модели (ModelPath). Нужен конктроль разработчика!",
                                    MessageType.Error);
                                continue;
                            }

                            #region Проверка РН, что он открыт, иначе - насильно открываю (только через создание элемента в проекте)
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

                                    List<ElementId> ids = new List<ElementId> { elementId };

                                    //This command will actualy open workset
                                    uidoc.ShowElements(ids);

                                    //Delete temporary cable tray
                                    doc.Delete(elementId);

                                    t.Commit();
                                }
                            }
                            #endregion

                            // Обновляю по новому пути
                            string resulModelPathString = ModelPathUtils.ConvertModelPathToUserVisiblePath(newModelPath);
                            linkType.LoadFrom(newModelPath, openConfig);
                            succsesSteps++;
                        }
                        catch (FileArgumentNotFoundException)
                        {
                            HtmlOutput.Print($"Файла по переопределенному пути нет. Проверь наличие файла для обновления тут: {ModelPathUtils.ConvertModelPathToUserVisiblePath(newModelPath)}. Если она там есть - обратись к разработчику!", MessageType.Error);
                        }
                    }
                }

                HtmlOutput.Print(
                    $"Успешно для {succsesSteps} моделей из {allModlesForChange} моделей, подлежащих замене пути",
                    MessageType.Success);
            }
            finally
            {
                commandData.Application.DialogBoxShowing -= OnDialogBoxShowing;
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Обработка события всплывающего окна Ревит
        /// </summary>=
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            if (args.Cancellable)
            {
                args.Cancel();
            }
            else
            {
                DBRevitDialog currentDBDialog = null;
                if (string.IsNullOrEmpty(args.DialogId))
                {
                    TaskDialogShowingEventArgs taskDialogShowingEventArgs = args as TaskDialogShowingEventArgs;
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => !string.IsNullOrEmpty(rd.Message) && taskDialogShowingEventArgs.Message.Contains(rd.Message));
                }
                else
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => args.DialogId.Contains(rd.DialogId));

                if (currentDBDialog == null)
                    HtmlOutput.Print($"Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека", MessageType.Error);

                if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                {
                    bool isOverride = args.OverrideResult((int)taskDialogResult);
                    if (!isOverride)
                        HtmlOutput.Print($"Окно {args.DialogId} не удалось обработать. Была применена команда {currentDBDialog.OverrideResult}, но она не сработала!", MessageType.Error);
                }
                else
                    HtmlOutput.Print($"Не удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!", MessageType.Error);
            }
        }

        /// <summary>
        /// Обработчик ошибок. Он нужен, когда закрывание окна не работает "Error dialog has no callback"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            FailuresAccessor fa = args.GetFailuresAccessor();
            IList<FailureMessageAccessor> failures = fa.GetFailureMessages();
            if (failures.Count > 0)
            {
                foreach (FailureMessageAccessor failure in failures)
                {
                    fa.DeleteWarning(failure);
                }
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

        private string RemoveSubstringIfExists(string originalString, string substring)
        {
            int index = originalString.IndexOf(substring);
            if (index != -1)
            {
                // Удаляем подстроку, если она найдена
                return originalString.Remove(index, substring.Length);
            }

            // Возвращаем исходную строку, если подстрока не найдена
            return originalString;
        }
    }
}
