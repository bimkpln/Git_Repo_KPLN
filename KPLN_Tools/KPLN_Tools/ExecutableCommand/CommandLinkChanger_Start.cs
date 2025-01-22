using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib.WorksetUtil;
using KPLN_Tools.Common.LinkManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandLinkChanger_Start : IExecutableCommand
    {
        internal protected static UIControlledApplication RevitUIControlledApp { get; set; }
        
        private static DBRevitDialog[] _dbRevitDialogs = null;

        private readonly LinkManagerEntity[] _linkChangeEntityColl;
        private readonly StringBuilder _sbErrResult = new StringBuilder();
        private readonly StringBuilder _sbWrnResult = new StringBuilder();
        private readonly StringBuilder _sbSuccResult = new StringBuilder();

        public CommandLinkChanger_Start(LinkManagerEntity[] linkChangeEntityColl)
        {
            _linkChangeEntityColl = linkChangeEntityColl;

            if (_dbRevitDialogs == null)
            {
                RevitDialogDbService currentRevitDialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
                _dbRevitDialogs = currentRevitDialogDbService.GetDBRevitDialogs().ToArray();
            }
        }


        internal static void SetStaticEnvironment(UIControlledApplication application)
        {
            RevitUIControlledApp = application;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;

            // Подписка на события
            app.DialogBoxShowing += OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.FailuresProcessing += OnFailureProcessing;

            if (_linkChangeEntityColl.Count() == 0)
                return Result.Cancelled;
            try
            {
                // Загрузка новых линков
                if (_linkChangeEntityColl.All(ent => ent is LinkManagerLoadEntity))
                    LoadNewLinks(doc);
                // Обновление существующих связей
                else if (_linkChangeEntityColl.All(ent => ent is LinkManagerUpdateEntity))
                    UpdateLinks(doc, uidoc);
                // Ошибка
                else
                {
                    HtmlOutput.Print($"Не удалось провести идентификацию алгоритма. Отправь разработчику", MessageType.Error);
                    return Result.Cancelled;
                }

                if (_sbErrResult.Length != 0)
                    HtmlOutput.Print($"Ошибки при загрузке связей связей: \n{_sbErrResult}", MessageType.Error);
                if (_sbWrnResult.Length != 0)
                    HtmlOutput.Print($"Нужно обратить внимание для связей: \n{_sbWrnResult}", MessageType.Warning);
                if (_sbSuccResult.Length != 0)
                    HtmlOutput.Print($"Успешный результат для связей: \n{_sbSuccResult}", MessageType.Success);
            }
            finally
            {
                app.DialogBoxShowing -= OnDialogBoxShowing;
                RevitUIControlledApp.ControlledApplication.FailuresProcessing -= OnFailureProcessing;
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

        /// <summary>
        /// Загрузка новых линков
        /// </summary>
        /// <param name="doc"></param>
        private void LoadNewLinks(Document doc)
        {
            // Коллекция RevitLinkInstance, для которых нужны отдельные РН
            List<RevitLinkInstance> instForWS = new List<RevitLinkInstance>();
            using (Transaction t = new Transaction(doc, $"KPLN: Загрузить связи"))
            {
                IEnumerable<string> rvtLinkTypeNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkType))
                    .Cast<RevitLinkType>()
                    .Select(rlt => rlt.Name);

                t.Start();

                foreach (LinkManagerLoadEntity linkLoadEntity in _linkChangeEntityColl.Cast<LinkManagerLoadEntity>())
                {
                    if (rvtLinkTypeNames.Any(rltn => rltn == linkLoadEntity.LinkName))
                        _sbWrnResult.AppendLine($"Связь '{linkLoadEntity.LinkName}' уже есть в проекте!\n");
                    else
                    {
                        ModelPath docModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkLoadEntity.LinkPath);

                        WorksetConfiguration wsConfig = CreatWSConfig(linkLoadEntity, docModelPath);
                        // Для РС не подходит относительный путь, а другой в API не представлен
                        bool isRelative = !linkLoadEntity.LinkPath.Contains("RSN");
                        RevitLinkOptions rlOpt;
                        if (wsConfig != null)
                            rlOpt = new RevitLinkOptions(isRelative, wsConfig);
                        else
                            rlOpt = new RevitLinkOptions(isRelative);

                        LinkLoadResult loadResult = null;
                        try
                        {
                            loadResult = RevitLinkType.Create(doc, docModelPath, rlOpt);
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            _sbErrResult.AppendLine($"Связь '{linkLoadEntity.LinkPath}' уже существует, или путь указан не верно");
                        }

                        if (loadResult != null)
                        {
                            // Успешно загружено
                            if (loadResult.LoadResult == LinkLoadResultType.LinkLoaded)
                            {
                                RevitLinkInstance linkInst = CreateLinkBySettings(doc, loadResult, linkLoadEntity);
                                if (linkLoadEntity.CreateWorksetForLinkInst)
                                    instForWS.Add(linkInst);
                            }
                            else if (loadResult.LoadResult == LinkLoadResultType.SameModelAsHost)
                                _sbWrnResult.AppendLine($"Связь '{linkLoadEntity.LinkPath}' нельзя загружать в такой же файл\n");
                            // Остальные статусы загрузок не совсем доходят, т.к. скорее сбрасывается exception (поэтому выше есть cath на ArgumentException), но разные типы зачем-то существуют
                            // Возможно будут изменения в будущем, или я не все учел
                            else
                                _sbErrResult.AppendLine($"Связь '{linkLoadEntity.LinkPath}' не обработана. Отправь разработчику\n");
                        }
                    }
                }

                t.Commit();
            }

            // Создание отдельного РН, если нужно.
            if (instForWS.Count() > 0)
                WorksetSetService.ExecuteFromService(doc, instForWS, null, false, false);
        }

        /// <summary>
        /// Обновление текущих линков
        /// </summary>
        /// <param name="doc"></param>
        private void UpdateLinks(Document doc, UIDocument uidoc)
        {
            foreach (LinkManagerUpdateEntity linkUpdateEntity in _linkChangeEntityColl.Cast<LinkManagerUpdateEntity>())
            {
                string oldModelPath = string.Empty;

                ModelPath linkNewModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkUpdateEntity.UpdatedLinkPath);
                WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);

                Element[] linkDocColl = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).ToArray();
                foreach (Element linkElem in linkDocColl)
                {
                    if (linkElem is RevitLinkType linkType)
                    {
                        try
                        {
                            if (!linkType.Name.ToLower().Equals($"{linkUpdateEntity.LinkName.ToLower()}.rvt"))
                                continue;
                        }
                        catch (InvalidObjectException ioe)
                        {
                            if (ioe.Message.Contains("he referenced object is not valid, possibly because it has been deleted from the database, or its creation was undone"))
                                continue;
                        }

                        // Анализирую положение связи 
                        ExternalFileReference extFileRef = linkType.GetExternalFileReference();
                        ModelPath oldExtPath = extFileRef.GetAbsolutePath();
                        oldModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(oldExtPath);

                        try
                        {
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
                                    Parameter wsparam = wall.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
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

                            linkType.LoadFrom(linkNewModelPath, openConfig);
                            _sbSuccResult.AppendLine($"Связь по пути {oldModelPath} усешно заменена на {linkUpdateEntity.UpdatedLinkPath}");
                        }
                        catch (FileArgumentNotFoundException)
                        {
                            _sbErrResult.AppendLine($"Файла по указанному пути нет. Проверь наличие файла для обновления тут: {linkUpdateEntity.UpdatedLinkPath}. Если он там есть - обратись к разработчику!");
                        }
                    }
                }
            }
        }

        private WorksetConfiguration CreatWSConfig(LinkManagerLoadEntity linkChangeEntity, ModelPath docModelPath)
        {
            try
            {
                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(docModelPath);
                IList<WorksetId> worksetIds = new List<WorksetId>();

                if (linkChangeEntity.WorksetToCloseNamesStartWith != null
                    && linkChangeEntity.WorksetToCloseNamesStartWith.Length > 0)
                {
                    string[] wsExceptions = linkChangeEntity.WorksetToCloseNamesStartWith.Split('~');
                    foreach (WorksetPreview worksetPrev in worksets)
                    {
                        if (linkChangeEntity.WorksetToCloseNamesStartWith.Count() == 0)
                            worksetIds.Add(worksetPrev.Id);
                        else if (!wsExceptions.Any(name => worksetPrev.Name.StartsWith(name)))
                            worksetIds.Add(worksetPrev.Id);
                    }
                    WorksetConfiguration wsConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                    wsConfig.Open(worksetIds);

                    return wsConfig;
                }
                else
                    return new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            }
            catch (CentralModelException)
            {
                return null;
            }
        }

        private RevitLinkInstance CreateLinkBySettings(Document doc, LinkLoadResult loadResult, LinkManagerLoadEntity linkChangeEntity)
        {
            string wsMsg;
            if (linkChangeEntity.CreateWorksetForLinkInst)
                wsMsg = $"Создан новый рабочий набор";
            else
                wsMsg = $"Пользователь отменил создание рабочего набора";

            string siteMsg;
            RevitLinkInstance result;
            // Создание instance
            try
            {
                result = RevitLinkInstance.Create(doc, loadResult.ElementId, linkChangeEntity.LinkCoordinateType.Type);

                siteMsg = "Выполнена загрузка по общей площадке";
                _sbSuccResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' загружена. Статус общей площадки: {siteMsg}. Статус создания нового рабочего набора: {wsMsg}.\n");
            }
            // Разные площадки
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                result = RevitLinkInstance.Create(doc, loadResult.ElementId, ImportPlacement.Origin);

                siteMsg = "Выполнена загрузка \"Совмещение внутренних начал\" (связь имеет разные площадки по сравнению с текущим файлов)";
                _sbWrnResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' загружена. Статус общей площадки: {siteMsg}. Статус создания нового рабочего набора: {wsMsg}.\n");
            }

            return result;
        }
    }
}
