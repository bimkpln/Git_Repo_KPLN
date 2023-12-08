using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;
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
        /// Список исключений в именах СЕМЕЙСТВ И ТИПОВ ревит (НАЧИНАЕТСЯ С) для генерации исключений в выбранных категориях.
        /// ВАЖНО: Имя семейства и типа в ревит прописано в параметре ELEM_FAMILY_AND_TYPE_PARAM, и формируется в формате "Имя семейства" + ": " + "Имя типа"
        /// </summary>
        private List<string> _exceptionFamilyAndTypeNameStartWithList;

        /// <summary>
        /// Список имен СЕМЕЙСТВ ревит для жёсткого поиска привязок
        /// </summary>
        private List<string> _hardCheckFamilyNamesList;

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

            #region Настройка под раздел
            string pathName = doc.PathName;
            if (pathName.ToUpper().Contains("_AR_")
                || pathName.ToUpper().Contains("_АР_")
                || (pathName.ToUpper().Contains("_AR.RVT") || pathName.ToUpper().Contains("_АР.RVT"))
                || (pathName.ToUpper().Contains("-AR.RVT") || pathName.ToUpper().Contains("-АР.RVT"))
                )
            {
                _exceptionFamilyAndTypeNameStartWithList = new List<string>()
                {
                    "199_",
                    "Базовая стена: 00_",
                };

                _hardCheckFamilyNamesList = new List<string>()
                {
                    "100_",
                    "115_",
                    "120_",
                };
            }
            else if (pathName.ToUpper().Contains("_KR_")
                || pathName.ToUpper().Contains("_КР_")
                || (pathName.ToUpper().Contains("_KR.RVT") || pathName.ToUpper().Contains("_КР.RVT"))
                || (pathName.ToUpper().Contains("-KR.RVT") || pathName.ToUpper().Contains("-КР.RVT"))
                )
            {
                // ЗАПОЛНИТЬ ДЛЯ КР
                _exceptionFamilyAndTypeNameStartWithList = new List<string>()
                {
                    "501_",
                };

                _hardCheckFamilyNamesList = new List<string>()
                {
                    "100_",
                    "115_",
                    "120_",
                };
            }
            else
                throw new Exception("Раздел проекта не определен. Обратись к разработчику!");
            #endregion

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

            List<LevelAndGridSolid> sectDataSolids = LevelAndGridSolid.PrepareSolids(doc, "КП_О_Секция");

            // 15 с
            Task.WaitAll(prepDataTask);

            // Первичный проход
            result.AddRange(CheckNullLevelElements(instDataColl));
            // 76 с 
            result.AddRange(CheckInstDataElems(instDataColl, sectDataSolids));

            // Повторный проход по НЕ проверенным
            result.AddRange(Reapeted_CheckElemsBySectionData(instDataColl, sectDataSolids));

            // Выдача ошибок
            result.Add(CheckNotAnalyzedElements(instDataColl));
            result.Add(ErrorGeomCheckedElementsResult());

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByCategory();
        }

        /// <summary>
        /// Получение элементов по списку категорий, с учетом фильтрации
        /// </summary>
        private Element[] PreapareElements(Document doc)
        {
            List<Element> result = new List<Element>();

            // Генерация фильтров
            List<FilterRule> filtRules = new List<FilterRule>();
            foreach (string currentName in _exceptionFamilyAndTypeNameStartWithList)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM), currentName, true);
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
                CheckStatus currentStatus;
                string approveComment = string.Empty;
                if (emptyLevelElems.All(e => ESEntity.ESBuilderUserText.IsDataExists_Text(e))
                    && emptyLevelElems.All(e =>
                        ESEntity.ESBuilderUserText.GetResMessage_Element(e).Description
                            .Equals(ESEntity.ESBuilderUserText.GetResMessage_Element(emptyLevelElems.FirstOrDefault()).Description)))
                {
                    currentStatus = CheckStatus.Approve;
                    approveComment = ESEntity.ESBuilderUserText.GetResMessage_Element(emptyLevelElems.FirstOrDefault()).Description;

                }
                else
                    currentStatus = CheckStatus.Error;

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
        /// Менеджер первичной проверки проверки коллекции элементов
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        /// <param name="sectDatas">Коллекция Solidов для проверки на пренадлежность к уровню/секции</param>
        /// <returns></returns>
        private IEnumerable<WPFEntity> CheckInstDataElems(CheckLevelOfInstanceData[] instDataColl, List<LevelAndGridSolid> sectDatas)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                //if (instData.CurrentElem.Id.IntegerValue != 29545897)
                //    continue;

                // Предварительная проверка на отсутсвие привязки выполнить ранее!
                if (instData.CurrentElemProjectDownLevel != null)
                {
                    WPFEntity checkExtraLevelBindElements = CheckExtraLevelBindElements(instData);
                    if (checkExtraLevelBindElements != null)
                        result.Add(checkExtraLevelBindElements);
                    
                    WPFEntity checkExtraOffsetstElements = CheckExtraOffsetstElements(instData);
                    if (checkExtraOffsetstElements != null)
                        result.Add(checkExtraOffsetstElements);

                    WPFEntity draftCheckElemsBySectionData = Draft_CheckElemsBySectionData(instData, sectDatas);
                    if (draftCheckElemsBySectionData != null)
                        result.Add(draftCheckElemsBySectionData);
                }
            }

            return result;
        }

        /// <summary>
        /// Проверка на нарушение привязки элементов с двумя уровнями
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <returns></returns>
        private WPFEntity CheckExtraLevelBindElements(CheckLevelOfInstanceData instData)
        {
            Level chkLvlDown = instData.CurrentElemProjectDownLevel;
            Level chkLvlUp = instData.CurrentElemProjectUpLevel;
            if (chkLvlDown != null && chkLvlUp != null)
            {
                string chkLvlDownNumber = LevelData.GetLevelNumber(chkLvlDown);
                string chkLvlUpNumber = LevelData.GetLevelNumber(chkLvlUp);
                if (int.TryParse(chkLvlDownNumber, out int chkDownNumber) && int.TryParse(chkLvlUpNumber, out int chkUpNumber))
                {
                    int lvlDiff = Math.Abs(Math.Abs(chkUpNumber) - Math.Abs(chkDownNumber));
                    if (lvlDiff > 1)
                    {
                        // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                        instData.IsEmptyChecked = false;
                        WPFEntity lvlDiffError = new WPFEntity(
                            instData.CurrentElem,
                            CheckStatus.Error,
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
        /// Проверка на наличие слишком большой привязки у элементов
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <returns></returns>
        private WPFEntity CheckExtraOffsetstElements(CheckLevelOfInstanceData instData)
        {
            string msg = string.Empty;
            if (Math.Abs(instData.DownOffset) > 20 && Math.Abs(instData.UpOffset) > 20)
                msg = $"Отсутп снизу составляет {instData.DownOffset}, оступ сверху состовляет {instData.UpOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            else if (Math.Abs(instData.DownOffset) > 20)
                msg = $"Отсутп снизу составляет {instData.DownOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            else if (Math.Abs(instData.UpOffset) > 20)
                msg = $"Отсутп сверху составляет {instData.UpOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            
            if (!string.IsNullOrEmpty(msg))
            {
                // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                instData.IsEmptyChecked = false;
                
                WPFEntity offsetError = new WPFEntity(
                    instData.CurrentElem,
                    CheckStatus.Error,
                    "Слишком большой оффсет элемента",
                    msg,
                    true,
                    true);
                offsetError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return offsetError;
            }

            return null;
        }

        /// <summary>
        /// Предварительная (элементы внутри секции) проверка элементов на принадлежность к секции и своему уровню
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <param name="sectDatas">Коллекция Solidов для проверки на пренадлежность к уровню/секции</param>
        private WPFEntity Draft_CheckElemsBySectionData(CheckLevelOfInstanceData instData, List<LevelAndGridSolid> sectDatas)
        {
            // Предварительная проверка на отсутсвие привязки выполнить ранее!
            if (instData.CurrentElemProjectDownLevel != null)
            {
                foreach (LevelAndGridSolid sectData in sectDatas)
                {
                    //if (instData.CurrentElem.Id.IntegerValue == 29545897)
                    //{
                    //    var a = 1;
                    //}
                    
                    // Игнорирую заведомо отличающиеся по отметкам секции
                    if (Math.Abs(instData.MinAndMaxElevation[0] - sectData.CurrentLevelData.MinAndMaxLvlPnts[0]) > 10
                        && Math.Abs(instData.MinAndMaxElevation[1] - sectData.CurrentLevelData.MinAndMaxLvlPnts[1]) > 10)
                        continue;

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
        private Solid GetIntesectedInstSolid(Solid instSolid, LevelAndGridSolid sectData)
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
        private WPFEntity CheckSolids(CheckLevelOfInstanceData instData, List<Solid> intersectSolids, LevelAndGridSolid sectData)
        {
            double instSolidArea = instData.CurrentSolidColl.Sum(ids => ids.SurfaceArea);
            double instSolidValue = instData.CurrentSolidColl.Sum(ids => ids.Volume);
            double intersectSolidsArea = intersectSolids?.Sum(intS => intS.SurfaceArea) ?? 0;
            double intersectSolidsValue = intersectSolids?.Sum(intS => intS.Volume) ?? 0;
            string instPrjLvlNumber = LevelData.GetLevelNumber(instData.CurrentElemProjectDownLevel);
            string sectCurrentLvlDownNumber = sectData.CurrentLevelData.CurrentDownLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentDownLevel)
                : string.Empty;
            string sectDataLvlCurrnetNumber = sectData.CurrentLevelData.CurrentLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentLevel)
                : string.Empty;
            string sectDataLvlAboveNumber = sectData.CurrentLevelData.CurrentAboveLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentAboveLevel)
                : string.Empty;

            // Анализ на смещение относительно уровня на 80% по однозначным элементам (жёсткая проверка) FamilyInstance и категориям (потолки, стены)
            if (((instData.CurrentElem is FamilyInstance familyInstance && _hardCheckFamilyNamesList.Where(hn => familyInstance.Symbol.FamilyName.Contains(hn)).Count() > 0) 
                    || instData.CurrentElem as Ceiling != null)
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.80
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                )
            {
                WPFEntity hardError = new WPFEntity(
                    instData.CurrentElem,
                    CheckStatus.Error,
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
                    CheckStatus.Error,
                    "Нарушена привязка к уровню, или уровню ниже",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя по факту расположен на уровне {sectData.CurrentLevelData.CurrentLevel.Name}",
                    true,
                    true,
                    "Перекрытия допускается привязывать либо к текущему уровню, либо к уровню ниже");
                hardFloorError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                return hardFloorError;
            }

            // Анализ на смещение стен
            else if ((instData.CurrentElem as Wall != null)
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                )
            {
                // Относительно уровня на более чем 1 для стен (УТОЧНИТЬ ПО ЭЛ-ТАМ КР!!!)
                if (sectData.CurrentLevelData.CurrentDownLevel != null
                    && sectData.CurrentLevelData.CurrentAboveLevel != null
                    && !instPrjLvlNumber.Equals(sectCurrentLvlDownNumber)
                    && !instPrjLvlNumber.Equals(sectDataLvlAboveNumber)
                    )
                {
                    WPFEntity hardError = new WPFEntity(
                        instData.CurrentElem,
                        CheckStatus.Error,
                        "Нарушены привязки к уровню ниже, или уровню выше",
                        $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                        true,
                        true);
                    hardError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                    return hardError;
                }

                // Относительно уровня на процент пересечения
                if (sectData.CurrentLevelData.CurrentDownLevel != null
                    && sectData.CurrentLevelData.CurrentAboveLevel != null
                    && instData.CurrentElemProjectDownLevel != null
                    && instData.CurrentElemProjectUpLevel != null
                    && intersectSolidsValue > instSolidValue * 0.80
                    )
                {
                    WPFEntity hardError = new WPFEntity(
                        instData.CurrentElem,
                        CheckStatus.Error,
                        "Нарушена привязка к уровню",
                        $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                        true,
                        true);
                    hardError.PrepareZoomGeometryExtension(instData.CurrentElem.get_BoundingBox(null));
                    return hardError;
                }

            }

            // Анализ на смещение относительно уровня на 50% по НЕ однозначным элементам
            else if (instData.CurrentElem as Floor == null
                && !(instData.CurrentElem is FamilyInstance famInst && _hardCheckFamilyNamesList.Where(hn => famInst.Symbol.FamilyName.Contains(hn)).Count() > 0)
                && instData.CurrentElem as Wall == null
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.50
                && !instPrjLvlNumber.Equals(sectDataLvlCurrnetNumber)
                )
            {
                WPFEntity warning = new WPFEntity(
                    instData.CurrentElem,
                    CheckStatus.Warning,
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
        private IEnumerable<WPFEntity> Reapeted_CheckElemsBySectionData(CheckLevelOfInstanceData[] instDataColl, List<LevelAndGridSolid> sectDatas)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                // Предварительная проверка на отсутсвие привязки выполнить ранее!
                if (instData.CurrentElemProjectDownLevel != null && instData.IsEmptyChecked)
                {

                    LevelAndGridSolid sectData = GetNearestSecData(instData, sectDatas);
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
        private LevelAndGridSolid GetNearestSecData(CheckLevelOfInstanceData instData, List<LevelAndGridSolid> sectDatas)
        {
            LevelAndGridSolid tempSect = null;
            double tempFacePrj = double.MaxValue;
            foreach (LevelAndGridSolid sectData in sectDatas)
            {
                double sectDataCurrentLvlElev = sectData.CurrentLevelData.CurrentLevel.Elevation;
                double instDataDownLvlElev = instData.CurrentElemProjectDownLevel.Elevation;
                if (sectDataCurrentLvlElev > instDataDownLvlElev + 30)
                    continue;
                
                FaceArray sectDataFaces = sectData.LevelSolid.Faces;
                foreach (XYZ instGeomCenter in instData.CurrentGeomCenterColl)
                {
                    foreach (Face face in sectDataFaces)
                    {
                        IntersectionResult prjRes = face.Project(instGeomCenter);
                        if (prjRes != null && prjRes.Distance < tempFacePrj)
                        {
                            tempSect = sectData;
                            tempFacePrj = prjRes.Distance;
                            break;
                        }
                    }
                }
            }

            return tempSect;
        }

        /// <summary>
        /// Проверка элементов, котоыре не прошли анализ ни одной из проверок
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        private WPFEntity CheckNotAnalyzedElements(CheckLevelOfInstanceData[] instDataColl) => new WPFEntity(
                instDataColl.Where(idc => idc.IsEmptyChecked).Select(idc => idc.CurrentElem),
                CheckStatus.Error,
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
            CheckStatus.Warning,
            "Элементы не удалось проверить из-за невозможности анализа геометрии",
            "Нужно проверить вручную",
            true,
            false
            );
    }
}
