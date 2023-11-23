using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using static KPLN_ModelChecker_User.Common.Collections;
namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckLevelOfInstances : AbstrCheckCommand<CommandCheckLevelOfInstances>, IExternalCommand
    {
        private readonly BuiltInCategory[] _catsToCheck = new BuiltInCategory[]
        {
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_GenericModel
        };

        /// <summary>
        /// Список исключений в именах семейств для генерации исключений в выбранных категориях
        /// </summary>
        private readonly List<string> _exceptionFamilyStartNameList = new List<string>
        {
            "199_",
            "501_",
        };

        /// <summary>
        /// Список имен семейств для жёсткого поиска привязок
        /// </summary>
        private readonly List<string> _hardCheckFamilyContainsName = new List<string>
        {
            "100_",
            "115_",
            "120_",
        };

        /// <summary>
        /// Коллекция элементов, которые не удалось проверить из-за ошибки с пересечением геометрии
        /// </summary>
        private List<Element> _errorGeomCheckedElements = new List<Element>();

        public CommandCheckLevelOfInstances() : base()
        {
        }

        internal CommandCheckLevelOfInstances(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, PreapareElements(doc));
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            #region Подготовливаю спец. класс отдельным потоком
            CheckLevelOfInstanceData[] instDataColl = null;
            Task prepDataTask = Task.Run(() =>
            {
                instDataColl = elemColl
                .Select(e =>
                    new CheckLevelOfInstanceData(e)
                    .SetCurrentSolidColl()
                    .SetCurrentBBoxColl()
                    .SetCurrentProjectLevel()
                    .SetOffsets()
                    )
                .ToArray();
            });
            #endregion

            #region Подготовливаю данные по делению проекта на секции и уровни
            List<CheckLevelOfInstanceGridData> gridDatas = CheckLevelOfInstanceGridData.GridPrepare(doc);
            List<CheckLevelOfInstanceLevelData> levelDatas = CheckLevelOfInstanceLevelData.LevelPrepare(doc);
            List<CheckLevelOfInstanceSectionData> sectData = CheckLevelOfInstanceSectionData.SpecialSolidPrepare(levelDatas, gridDatas);
            #endregion

            Task.WaitAll(prepDataTask);

            // Первичный проход
            result.AddRange(CheckNullLevelElements(instDataColl));
            result.AddRange(CheckInstDataElems(instDataColl, sectData));

            // Повторный проход по НЕ проверенным
            result.AddRange(Reapeted_CheckElemsBySectionData(instDataColl, sectData));
            result.Add(Reapeted_CheckNotAnalyzedElements(instDataColl));

            // Выдача ошибок
            result.Add(ErrorGeomCheckedElementsResult());

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Получение элементов по списку категорий, с учетом фильтрации
        /// </summary>
        private Element[] PreapareElements(Document doc)
        {
            List<Element> result = new List<Element>();

            // Генерация фильтров
            List<FilterRule> filtRules = new List<FilterRule>(_exceptionFamilyStartNameList.Count);
            foreach (string currentName in _exceptionFamilyStartNameList)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
                filtRules.Add(fRule);
            }
            ElementParameterFilter eFilter = new ElementParameterFilter(filtRules);

            // Генерация и фильтрация FilteredElementCollector
            List<FilteredElementCollector> bicColl = new List<FilteredElementCollector>(_catsToCheck.Count());
            foreach (BuiltInCategory bic in _catsToCheck)
            {
                FilteredElementCollector fic = new FilteredElementCollector(doc).OfCategory(bic);
                fic.WherePasses(eFilter).WhereElementIsNotElementType();
                bicColl.Add(fic);
            }

            // Добавляю очищенные элементы в коллекцию
            foreach (FilteredElementCollector coll in bicColl)
            {
                result.AddRange(coll
                    .Where(e => (e is FamilyInstance f && f.SuperComponent == null) || !(e is FamilyInstance)));
            }

            return result.ToArray();
        }


        /// <summary>
        /// Проверка элементов, у которых нет данных по уровню
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        private IEnumerable<WPFEntity> CheckNullLevelElements(CheckLevelOfInstanceData[] instDataColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            List<Element> emptyLevelElems = new List<Element>();
            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                if (instData.CurrentElemProjectDownLevel == null)
                {
                    // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                    instData.IsEmptyChecked = false;
                    emptyLevelElems.Add(instData.CurrentElem);
                }
            }

            if (emptyLevelElems.Count > 0)
            {
                Status currentStatus;
                string approveComment = string.Empty;
                if (emptyLevelElems.All(e => ESEntity.ESBuilderUserText.IsDataExists_Text(e))
                    && emptyLevelElems.All(e =>
                        ESEntity.ESBuilderUserText.GetResMessage_Element(e).Description
                            .Equals(ESEntity.ESBuilderUserText.GetResMessage_Element(emptyLevelElems.FirstOrDefault()).Description)))
                {
                    currentStatus = Status.Approve;
                    approveComment = ESEntity.ESBuilderUserText.GetResMessage_Element(emptyLevelElems.FirstOrDefault()).Description;

                }
                else
                    currentStatus = Status.Error;

                result.Add(new WPFEntity(
                    emptyLevelElems,
                    currentStatus,
                    "Уровень не заполнен",
                    $"У элементов не заполнен параметр уровня",
                    true,
                    true,
                    null,
                    approveComment
                    ));
            }

            return result;
        }

        /// <summary>
        /// Менеджер проверки коллекции элементов
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        /// <param name="sectDatas">Коллекция Solidов для проверки на пренадлежность к уровню/секции</param>
        /// <returns></returns>
        private IEnumerable<WPFEntity> CheckInstDataElems(CheckLevelOfInstanceData[] instDataColl, List<CheckLevelOfInstanceSectionData> sectDatas)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                // Предварительная проверка на отсутсвие привязки выполнить ранее!
                if (instData.CurrentElemProjectDownLevel != null)
                {
                    WPFEntity checkExtraLevelBindElements = CheckExtraLevelBindElements(instData);
                    if (checkExtraLevelBindElements != null)
                        result.Add(checkExtraLevelBindElements);
                    else
                    {
                        WPFEntity draftCheckElemsBySectionData = Draft_CheckElemsBySectionData(instData, sectDatas);
                        if (draftCheckElemsBySectionData != null)
                            result.Add(draftCheckElemsBySectionData);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Проверка на привязку для элементов с двумя уровнями
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <returns></returns>
        private WPFEntity CheckExtraLevelBindElements(CheckLevelOfInstanceData instData)
        {

            Level chkLvlDown = instData.CurrentElemProjectDownLevel;
            Level chkLvlUp = instData.CurrentElemProjectUpLevel;
            if (chkLvlDown != null && chkLvlUp != null)
            {
                string chkLvlDownNumber = CheckLevelOfInstanceLevelData.GetLevelNumber(chkLvlDown);
                string chkLvlUpNumber = CheckLevelOfInstanceLevelData.GetLevelNumber(chkLvlUp);
                if (int.TryParse(chkLvlDownNumber, out int chkDownNumber) && int.TryParse(chkLvlUpNumber, out int chkUpNumber))
                {
                    int lvlDiff = Math.Abs(Math.Abs(chkUpNumber) - Math.Abs(chkDownNumber));
                    if (lvlDiff > 1)
                    {
                        // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                        instData.IsEmptyChecked = false;
                        WPFEntity lvlDiffError = new WPFEntity(
                            instData.CurrentElem,
                            Status.Error,
                            "Нарушено деление по уровням",
                            $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name} снизу и к {instData.CurrentElemProjectUpLevel.Name} сверху. Запрещено моделировать элементы без деления на уровни",
                            true,
                            true);
                        lvlDiffError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                        return lvlDiffError;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Предварительная (элементы внутри секции) проверка элементов на принадлежность к секции и своему уровню
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <param name="sectDatas">Коллекция Solidов для проверки на пренадлежность к уровню/секции</param>
        private WPFEntity Draft_CheckElemsBySectionData(CheckLevelOfInstanceData instData, List<CheckLevelOfInstanceSectionData> sectDatas)
        {
            // Предварительная проверка на отсутсвие привязки выполнить ранее!
            if (instData.CurrentElemProjectDownLevel != null)
            {
                //if (instData.CurrentElem.Id.IntegerValue != 7218226)
                //    continue;

                foreach (CheckLevelOfInstanceSectionData sectData in sectDatas)
                {
                    List<Solid> intersectSolids = new List<Solid>();
                    try
                    {
                        foreach (Solid instSolid in instData.CurrentSolidColl)
                        {
                            if (instSolid.Volume == 0)
                                continue;

                            // Првоеряю положение в секции
                            Solid checkSectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(instSolid, sectData.LevelSolid, BooleanOperationsType.Intersect);
                            if (checkSectSolid != null && checkSectSolid.Volume > 0)
                            {
                                Solid resSolid = GetIntesectedInstSolid(instSolid, sectData);
                                if (resSolid != null)
                                    intersectSolids.Add(resSolid);
                            }
                        }

                        if (intersectSolids.Count != 0)
                        {
                            // Выставляю флаг, что элемент принадлежит хоть какой-то секции, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                            instData.IsEmptyChecked = false;
                            WPFEntity checkSolids = CheckSolids(instData, intersectSolids, sectData);
                            if (checkSolids != null)
                                return checkSolids;
                        }
                    }
                    // Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        instData.IsEmptyChecked = false;
                        _errorGeomCheckedElements.Add(instData.CurrentElem);
                        return null;
                    }
                    catch (Exception ex)
                    {
                        Print($"Первичная проверка: Что-то непонятное с элементом с id: {instData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}", MessageType.Error);
                        return null;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Определение солида, который пересекается с солидом секции. Солид эл-та Ревит ПРИТЯГИВАЕТСЯ к солиду секции
        /// </summary>
        /// <param name="instSolid">Солид эл-та ревит для проверки</param>
        /// <param name="sectData">Солид секции для проверки</param>
        /// <returns></returns>
        private Solid GetIntesectedInstSolid(Solid instSolid, CheckLevelOfInstanceSectionData sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.LevelSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(sectTransform.Origin.X, sectTransform.Origin.Y, instTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.LevelSolid, BooleanOperationsType.Intersect);
            if (intersectSolid != null && intersectSolid.Volume > 0)
                return intersectSolid;

            return null;
        }

        /// <summary>
        /// Проанализировать CheckLevelOfInstanceData и его коллекцию солидов, пересекающихся с солидом CheckLevelOfInstanceSectionData 
        /// </summary>
        /// <param name="instData">Спец. класс эл-та Ревит</param>
        /// <param name="intersectSolids">Список солидов, эл-та ревит, которые пересекаются с солидом спец. класса секции Ревит</param>
        /// <param name="sectData">Спец. класс секции Ревит</param>
        /// <returns></returns>
        private WPFEntity CheckSolids(CheckLevelOfInstanceData instData, List<Solid> intersectSolids, CheckLevelOfInstanceSectionData sectData)
        {
            double instSolidArea = instData.CurrentSolidColl.Sum(ids => ids.SurfaceArea);
            double instSolidValue = instData.CurrentSolidColl.Sum(ids => ids.Volume);
            double intersectSolidsArea = intersectSolids?.Sum(intS => intS.SurfaceArea) ?? 0;
            double intersectSolidsValue = intersectSolids?.Sum(intS => intS.Volume) ?? 0;
            string instPrjLvlNumber = CheckLevelOfInstanceLevelData.GetLevelNumber(instData.CurrentElemProjectDownLevel);
            string sectCurrentLvlDownNumber = sectData.CurrentLevelData.CurrentDownLevel != null
                ? CheckLevelOfInstanceLevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentDownLevel)
                : string.Empty;
            string sectDataLvlCurrnetNumber = sectData.CurrentLevelData.CurrentLevel != null
                ? CheckLevelOfInstanceLevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentLevel)
                : string.Empty;
            string sectDataLvlAboveNumber = sectData.CurrentLevelData.CurrentAboveLevel != null
                ? CheckLevelOfInstanceLevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentAboveLevel)
                : string.Empty;

            // Анализ на смещение относительно уровня на 80% по однозначным элементам (жёсткая проверка) FamilyInstance и категориям (потолки)
            if (((instData.CurrentElem is FamilyInstance familyInstance && _hardCheckFamilyContainsName.Where(hn => familyInstance.Symbol.FamilyName.Contains(hn)).Count() > 0) 
                    || instData.CurrentElem as Ceiling != null)
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.80
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                )
            {
                WPFEntity hardError = new WPFEntity(
                    instData.CurrentElem,
                    Status.Error,
                    "Нарушена привязка к уровню",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                    true,
                    true,
                    "Элемент попадает в список для постоянной привязки к текущему уровню");
                hardError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return hardError;
            }

            // Анализ на смещение относительно уровня на 1 уровень ниже, или на текущем уровне - ТОЛЬКО для перекрытий
            else if (instData.CurrentElem as Floor != null
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && sectData.CurrentLevelData.CurrentAboveLevel != null
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                && !instPrjLvlNumber.Equals(sectDataLvlAboveNumber)
                )
            {
                WPFEntity hardFloorError = new WPFEntity(
                    instData.CurrentElem,
                    Status.Error,
                    "Нарушена привязка к уровню, или уровню ниже",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя по факту расположен на уровне {sectData.CurrentLevelData.CurrentLevel.Name}",
                    true,
                    true,
                    "Перекрытия допускается привязывать либо к текущему уровню, либо к уровню ниже");
                hardFloorError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return hardFloorError;
            }

            // Анализ на смещение относительно уровня на более чем 1 для стен (УТОЧНИТЬ ПО ЭЛ-ТАМ КР!!!)
            else if ((instData.CurrentElem as Wall != null)
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && sectData.CurrentLevelData.CurrentAboveLevel != null
                && !instPrjLvlNumber.Equals(sectCurrentLvlDownNumber)
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                && !instPrjLvlNumber.Equals(sectDataLvlAboveNumber)
                )
            {
                WPFEntity hardError = new WPFEntity(
                    instData.CurrentElem,
                    Status.Error,
                    "Нарушены привязки к уровню ниже, или уровню выше",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                    true,
                    true);
                hardError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return hardError;
            }

            // Анализ на смещение относительно уровня на 50% по НЕ однозначным элементам
            else if (instData.CurrentElem as Floor == null
                && !(instData.CurrentElem is FamilyInstance famInst && _hardCheckFamilyContainsName.Where(hn => famInst.Symbol.FamilyName.Contains(hn)).Count() > 0)
                && instData.CurrentElem as Wall == null
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.50
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                )
            {
                WPFEntity warning = new WPFEntity(
                    instData.CurrentElem,
                    Status.Warning,
                    "Необходим контроль",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                    true,
                    false,
                    "Устранять ошибки не обязательно, но проверить положение элементов стоит");
                warning.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return warning;
            }

            return null;
        }

        /// <summary>
        /// Повторная (элементы за пределами секции) проверка элементов на принадлежность к секции
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        /// <param name="sectDatas">Коллекция Solidов для проверки на пренадлежность к уровню/секции</param>
        private IEnumerable<WPFEntity> Reapeted_CheckElemsBySectionData(CheckLevelOfInstanceData[] instDataColl, List<CheckLevelOfInstanceSectionData> sectDatas)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                if (instData.CurrentElem.Id.IntegerValue == 29835789)
                {
                    var a = 1;
                }

                if (instData.CurrentElem.Id.IntegerValue == 7218226)
                {
                    var aф = 1;
                }
                // Предварительная проверка на отсутсвие привязки выполнить ранее!
                if (instData.CurrentElemProjectDownLevel != null && instData.IsEmptyChecked)
                {

                    CheckLevelOfInstanceSectionData sectData = GetNearestSecData(instData, sectDatas);
                    if (sectData != null)
                    {
                        try
                        {
                            List<Solid> intersectSolids = new List<Solid>();
                            foreach (Solid instSolid in instData.CurrentSolidColl)
                            {
                                if (instSolid.Volume == 0)
                                    continue;

                                Solid resSolid = GetIntesectedInstSolid(instSolid, sectData);
                                if (resSolid != null)
                                    intersectSolids.Add(resSolid);
                            }

                            if (intersectSolids.Count == 0)
                                continue;
                            else
                            {
                                // Выставляю флаг, что элемент принадлежит хоть какой-то секции, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                                instData.IsEmptyChecked = false;
                                WPFEntity checkSolids = CheckSolids(instData, intersectSolids, sectData);
                                if (checkSolids != null)
                                    result.Add(checkSolids);
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            instData.IsEmptyChecked = false;
                            _errorGeomCheckedElements.Add(instData.CurrentElem);
                        }
                        catch (Exception ex)
                        {
                            Print($"Что-то непонятное с элементом с id: {instData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}", MessageType.Error);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Получить ближайшую подходящую секцию для текущего CheckLevelOfInstanceData
        /// </summary>
        /// <param name="instData">Спец. класс для проверки</param>
        /// <returns></returns>
        private CheckLevelOfInstanceSectionData GetNearestSecData(CheckLevelOfInstanceData instData, List<CheckLevelOfInstanceSectionData> sectDatas)
        {
            CheckLevelOfInstanceSectionData tempSect = null;
            double tempFacePrj = double.MaxValue;
            foreach (CheckLevelOfInstanceSectionData sectData in sectDatas)
            {
                FaceArray sectDataFaces = sectData.LevelSolid.Faces;
                foreach (Solid instSolid in instData.CurrentSolidColl)
                {
                    if (instSolid.Volume != 0)
                    {
                        XYZ instCentr = instSolid.ComputeCentroid();
                        foreach (Face face in sectDataFaces)
                        {
                            IntersectionResult prjRes = face.Project(instCentr);
                            if (prjRes != null && prjRes.Distance < tempFacePrj)
                                tempSect = sectData;
                        }
                    }
                }
            }

            return tempSect;
        }

        /// <summary>
        /// Повторная проверка элементов, котоыре не прошли анализ ни одной из проверок
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        private WPFEntity Reapeted_CheckNotAnalyzedElements(CheckLevelOfInstanceData[] instDataColl) => new WPFEntity(
                instDataColl.Where(idc => idc.IsEmptyChecked).Select(idc => idc.CurrentElem),
                Status.Error,
                "Элементы не удалось проверить из-за особенностей проекта",
                $"Необходимо показать разработчику",
                true,
                false);

        /// <summary>
        /// Преобразование в WPFEntity коллеции с ошибками при анализе геометрии
        /// </summary>
        /// <returns></returns>
        private WPFEntity ErrorGeomCheckedElementsResult() => new WPFEntity(
            _errorGeomCheckedElements,
            Status.Warning,
            "Элементы не удалось проверить из-за невозможности анализа геометрии",
            "Нужно проверить вручную",
            true,
            false
            );
    }
}
