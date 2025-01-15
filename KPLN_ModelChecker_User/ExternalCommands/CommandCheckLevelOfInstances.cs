using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
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

        private string _sectParamName;
        private string _lvlIndexParamName;

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

            _sectParamName = "КП_О_Секция";
            _lvlIndexParamName = "КП_О_Этаж";
            if (doc.Title.StartsWith("СЕТ_1"))
            {
                _sectParamName = "СМ_Секция";
                _lvlIndexParamName = "СМ_Этаж";
            }
            List<LevelAndGridSolid> sectDataSolids = LevelAndGridSolid.PrepareSolids(doc, _sectParamName, _lvlIndexParamName);

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
                result.Add(new WPFEntity(
                    ESEntity,
                    emptyLevelElems,
                    "Уровень не заполнен",
                    $"У элементов не заполнен параметр уровня",
                    string.Empty,
                    false));
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
                string chkLvlDownNumber = LevelData.GetLevelNumber(chkLvlDown, _lvlIndexParamName);
                string chkLvlUpNumber = LevelData.GetLevelNumber(chkLvlUp, _lvlIndexParamName);
                
                if (int.TryParse(chkLvlDownNumber, out int chkDownNumber) && int.TryParse(chkLvlUpNumber, out int chkUpNumber))
                {
                    int lvlDiff = Math.Abs(Math.Abs(chkUpNumber) - Math.Abs(chkDownNumber));
                    if (lvlDiff > 1)
                    {
                        // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                        instData.IsEmptyChecked = false;

                        WPFEntity lvlDiffError = new WPFEntity(
                            ESEntity,
                            instData.CurrentElem,
                            "Нарушено деление по уровням",
                            $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name} снизу и к {instData.CurrentElemProjectUpLevel.Name} сверху. Запрещено моделировать элементы без деления на уровни",
                            string.Empty,
                            true,
                            true);

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
                msg = $"Отсуп снизу составляет {instData.DownOffset}, отступ сверху составляет {instData.UpOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            else if (Math.Abs(instData.DownOffset) > 20)
                msg = $"Отсуп снизу составляет {instData.DownOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            else if (Math.Abs(instData.UpOffset) > 20)
                msg = $"Отсуп сверху составляет {instData.UpOffset}. Запрещено моделировать элементы без корректного назначения уровня";
            
            if (!string.IsNullOrEmpty(msg))
            {
                // Выставляю флаг, что элемент не прошел текущую проверку, чтобы ниже зацепить эл-ты, которые не подверглись проверкам
                instData.IsEmptyChecked = false;

                WPFEntity offsetError = new WPFEntity(
                    ESEntity,
                    instData.CurrentElem,
                    "Слишком большой оффсет элемента",
                    msg,
                    string.Empty,
                    true,
                    true);

                return offsetError;
            }

            return null;
        }

        /// <summary>
        /// Предварительная (элементы внутри секции) проверка элементов на принадлежность к секции и своему уровню
        /// </summary>
        /// <param name="instData">Элемент на проверку</param>
        /// <param name="sectDatas">Коллекция солидов для проверки на принадлежность к уровню/секции</param>
        private WPFEntity Draft_CheckElemsBySectionData(CheckLevelOfInstanceData instData, List<LevelAndGridSolid> sectDatas)
        {
            // Предварительная проверка на отсутствие привязки выполнить ранее!
            if (instData.CurrentElemProjectDownLevel == null) 
                return null;
            
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

                        // Проверяю положение в секции
                        Solid checkSectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(instSolid, sectData.CurrentSolid, BooleanOperationsType.Intersect);
                        if (checkSectSolid == null || !(checkSectSolid.Volume > 0)) 
                            continue;
                            
                        Solid resSolid = GetIntesectedInstSolid(instSolid, sectData);
                        if (resSolid != null)
                            intersectSolids.Add(resSolid);
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
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета (проблемы с элементами "по касательной")
            Transform sectTransform = sectData.CurrentSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(sectTransform.Origin.X, sectTransform.Origin.Y, instTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.CurrentSolid, BooleanOperationsType.Intersect);
            if (intersectSolid != null && intersectSolid.Volume > 0)
                return intersectSolid;

            return null;
        }

        /// <summary>
        /// Проанализировать CheckLevelOfInstanceData и его коллекцию солидов, пересекающихся с солидом CheckLevelOfInstanceSectionData 
        /// </summary>
        /// <param name="instData">Спец. Класс эл-та Ревит</param>
        /// <param name="intersectSolids">Список солидов, эл-та ревит, которые пересекаются с солидом спец. Класса секции Ревит</param>
        /// <param name="sectData">Спец. Класс секции Ревит</param>
        /// <returns></returns>
        private WPFEntity CheckSolids(CheckLevelOfInstanceData instData, List<Solid> intersectSolids, LevelAndGridSolid sectData)
        {
            double instSolidArea = instData.CurrentSolidColl.Sum(ids => ids.SurfaceArea);
            double instSolidValue = instData.CurrentSolidColl.Sum(ids => ids.Volume);
            double intersectSolidsArea = intersectSolids?.Sum(intS => intS.SurfaceArea) ?? 0;
            double intersectSolidsValue = intersectSolids?.Sum(intS => intS.Volume) ?? 0;
            string instPrjLvlNumber = LevelData.GetLevelNumber(instData.CurrentElemProjectDownLevel, _lvlIndexParamName);
            string sectCurrentLvlDownNumber = sectData.CurrentLevelData.CurrentDownLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentDownLevel, _lvlIndexParamName)
                : string.Empty;
            string sectDataLvlCurrentNumber = sectData.CurrentLevelData.CurrentLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentLevel, _lvlIndexParamName)
                : string.Empty;
            string sectDataLvlAboveNumber = sectData.CurrentLevelData.CurrentAboveLevel != null
                ? LevelData.GetLevelNumber(sectData.CurrentLevelData.CurrentAboveLevel, _lvlIndexParamName)
                : string.Empty;

            // Анализ на смещение относительно уровня на 80% по однозначным элементам (жёсткая проверка) FamilyInstance и категориям (потолки, стены)
            if (((instData.CurrentElem is FamilyInstance familyInstance 
                  && _hardCheckFamilyNamesList.Any(hn => familyInstance.Symbol.FamilyName.Contains(hn))) 
                    || instData.CurrentElem is Ceiling)
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.80
                && !instPrjLvlNumber.Equals(sectDataLvlCurrentNumber)
                )
            {
                WPFEntity hardError = new WPFEntity(
                    ESEntity,
                    instData.CurrentElem,
                    "Нарушена привязка к уровню",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                    "Элемент попадает в список для постоянной привязки к текущему уровню",
                    true,
                    true);

                return hardError;
            }

            // Анализ на смещение относительно уровня на 1 уровень ниже, или на текущем уровне - ТОЛЬКО для перекрытий
            else if (instData.CurrentElem is Floor
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && sectData.CurrentLevelData.CurrentAboveLevel != null
                && !instPrjLvlNumber.Equals(sectDataLvlCurrentNumber)
                && !instPrjLvlNumber.Equals(sectDataLvlAboveNumber)
                )
            {
                WPFEntity hardFloorError = new WPFEntity(
                    ESEntity,
                    instData.CurrentElem,
                    "Нарушена привязка к уровню, или уровню ниже",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя по факту расположен на уровне {sectData.CurrentLevelData.CurrentLevel.Name}",
                    "Перекрытия допускается привязывать либо к текущему уровню, либо к уровню ниже",
                    true,
                    true);

                return hardFloorError;
            }

            // Анализ на смещение стен
            else if ((instData.CurrentElem is Wall)
                && !instPrjLvlNumber.Equals(sectDataLvlCurrentNumber)
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
                        ESEntity,
                        instData.CurrentElem,
                        "Нарушены привязки к уровню ниже, или уровню выше",
                        $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                        string.Empty,
                        true,
                        true);

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
                        ESEntity,
                        instData.CurrentElem,
                        "Нарушена привязка к уровню",
                        $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                        string.Empty,
                        true,
                        true);

                    return hardError;
                }

            }

            // Анализ на смещение относительно уровня на 50% по НЕ однозначным элементам
            else if (!(instData.CurrentElem is Floor)
                && !(instData.CurrentElem is FamilyInstance famInst && _hardCheckFamilyNamesList.Any(hn => famInst.Symbol.FamilyName.Contains(hn)))
                && !(instData.CurrentElem is Wall)
                && sectData.CurrentLevelData.CurrentDownLevel != null
                && intersectSolidsValue > instSolidValue * 0.50
                && !instPrjLvlNumber.Equals(sectDataLvlCurrentNumber)
                )
            {
                WPFEntity warning = new WPFEntity(
                    ESEntity,
                    instData.CurrentElem,
                    "Необходим контроль",
                    $"Элемент привязан к {instData.CurrentElemProjectDownLevel.Name}, хотя на {Math.Round(intersectSolidsValue / instSolidValue, 2) * 100}% подходит уровню {sectData.CurrentLevelData.CurrentLevel.Name}",
                    "Устранять ошибки не обязательно, но проверить положение элементов стоит",
                    true,
                    false);

                return warning;
            }

            return null;
        }

        /// <summary>
        /// Повторная (элементы за пределами секции) проверка элементов на принадлежность к секции
        /// </summary>
        /// <param name="instDataColl">Коллекция на проверку</param>
        /// <param name="sectDatas">Коллекция солидов для проверки на принадлежность к уровню/секции</param>
        private IEnumerable<WPFEntity> Reapeted_CheckElemsBySectionData(CheckLevelOfInstanceData[] instDataColl, List<LevelAndGridSolid> sectDatas)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckLevelOfInstanceData instData in instDataColl)
            {
                // Предварительная проверка на отсутствие привязки выполнить ранее!
                if (instData.CurrentElemProjectDownLevel == null || !instData.IsEmptyChecked) 
                    continue;
                
                LevelAndGridSolid sectData = GetNearestSecData(instData, sectDatas);
                if (sectData == null) 
                    continue;
                
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
                
                FaceArray sectDataFaces = sectData.CurrentSolid.Faces;
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
                ESEntity,
                instDataColl.Where(idc => idc.IsEmptyChecked).Select(idc => idc.CurrentElem),
                "Элементы не удалось проверить из-за особенностей проекта",
                $"Необходимо показать разработчику",
                string.Empty,
                false);

        /// <summary>
        /// Преобразование в WPFEntity коллеции с ошибками при анализе геометрии
        /// </summary>
        /// <returns></returns>
        private WPFEntity ErrorGeomCheckedElementsResult() => new WPFEntity(
            ESEntity,
            _errorGeomCheckedElements,
            "Элементы не удалось проверить из-за невозможности анализа геометрии",
            "Нужно проверить вручную",
            string.Empty,
            false);
    }
}
