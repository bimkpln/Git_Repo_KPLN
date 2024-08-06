using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib.WorksetUtil;
using KPLN_Tools.Common.LinkManager;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandLinkChanger_Start : IExecutableCommand
    {
        private readonly LinkManagerEntity[] _linkChangeEntityColl;
        private readonly StringBuilder _sbErrResult = new StringBuilder();
        private readonly StringBuilder _sbWrnResult = new StringBuilder();
        private readonly StringBuilder _sbSuccResult = new StringBuilder();

        public CommandLinkChanger_Start(LinkManagerEntity[] linkChangeEntityColl)
        {
            _linkChangeEntityColl = linkChangeEntityColl;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            // Коллекция RevitLinkInstance, для которых нужны отдельные РН
            List<RevitLinkInstance> instForWS = new List<RevitLinkInstance>();
            using (Transaction t = new Transaction(doc, $"KPLN: Связи rvt"))
            {
                t.Start();

                foreach (LinkManagerEntity linkChangeEntity in _linkChangeEntityColl)
                {
                    ModelPath docModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkChangeEntity.LinkPath);

                    WorksetConfiguration wsConfig = CreatWSConfig(linkChangeEntity, docModelPath);
                    // Для РС не подходит относительный путь, а другой в API не представлен
                    bool isRelative = !linkChangeEntity.LinkPath.Contains("RSN");
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
                    catch (ArgumentException)
                    {
                        _sbErrResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' уже существует, или путь указан не верно");
                    }

                    if (loadResult != null)
                    {
                        // Успешно загружено
                        if (loadResult.LoadResult == LinkLoadResultType.LinkLoaded)
                        {
                            RevitLinkInstance linkInst = CreateLinkBySettings(doc, loadResult, linkChangeEntity);
                            if (linkChangeEntity.CreateWorksetForLinkInst)
                                instForWS.Add(linkInst);
                        }
                        else if (loadResult.LoadResult == LinkLoadResultType.SameModelAsHost)
                            _sbWrnResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' нельзя загружать в такой же файл\n");
                        // Остальные статусы загрузок не совсем доходят, т.к. скорее сбрасывается exception (поэтому выше есть cath на ArgumentException), но разные типы зачем-то существуют
                        // Возможно будут изменения в будущем, или я не все учел
                        else
                            _sbErrResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' не обработана. Отправь разработчику\n");
                    }
                }

                t.Commit();
            }

            // Создание отдельного РН, если нужно.
            if (instForWS.Count() > 0)
                WorksetSetService.ExecuteFromService(doc, instForWS, false);

            if (_sbErrResult.Length != 0)
                HtmlOutput.Print($"Ошибки при загрузке связей связей: \n{_sbErrResult}", MessageType.Error);
            if (_sbWrnResult.Length != 0)
                HtmlOutput.Print($"Нужно обратить внимание для связей: \n{_sbWrnResult}", MessageType.Warning);
            if (_sbSuccResult.Length != 0)
                HtmlOutput.Print($"Успешный результат для связей: \n{_sbSuccResult}", MessageType.Success);

            return Result.Succeeded;
        }

        private WorksetConfiguration CreatWSConfig(LinkManagerEntity linkChangeEntity, ModelPath docModelPath)
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

        private RevitLinkInstance CreateLinkBySettings(Document doc, LinkLoadResult loadResult, LinkManagerEntity linkChangeEntity)
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
            catch (InvalidOperationException)
            {
                result = RevitLinkInstance.Create(doc, loadResult.ElementId, ImportPlacement.Origin);

                siteMsg = "Выполнена загрузка \"Совмещение внутренних начал\" (связь имеет разные площадки по сравнению с текущим файлов)";
                _sbWrnResult.AppendLine($"Связь '{linkChangeEntity.LinkPath}' загружена. Статус общей площадки: {siteMsg}. Статус создания нового рабочего набора: {wsMsg}.\n");
            }

            return result;
        }
    }
}
