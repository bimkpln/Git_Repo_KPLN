using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Controls;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckMEPHeight : AbstrCheckCommand<CommandCheckMEPHeight>, IExternalCommand
    {
        /// <summary>
        /// Список категорий элементов для проверки
        /// </summary>
        private readonly List<BuiltInCategory> _mepBICColl = new List<BuiltInCategory>()
        {
            // Системные семейства
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctCurvesInsulation,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeInsulations,
            BuiltInCategory.OST_CableTray,
            // Пользовательские семейства
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_PlumbingFixtures,
        };

        /// <summary>
        /// Список исключений в именах семейств для генерации исключений в выбранных категориях
        /// </summary>
        private readonly List<string> _exceptionFamilyStartNameList = new List<string>
        {
            "490_",
            "491_",
            "500_",
            "501_",
            "502_",
            "503_",
            "708_",
            "709_",
            "710_",
            "711_",
            "712_",
            "715_",
            "720_",
            "810_",
            "910_",
            "915_",
            "920_",
            "951_",
            "952_",
            "953_",
            "960_",
        };

        public CommandCheckMEPHeight() : base()
        {
        }

        internal CommandCheckMEPHeight(ExtensibleStorageEntity esEntity) : base(esEntity)
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
            WPFEntity[] wpfColl = CheckCommandRunner(doc, PreapareIOSElements(doc));
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) 
                form.Show(); 
            else 
                return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl)
        {
            if (!objColl.Any())
                throw new CheckerException("В проекте отсутсвуют необходимые элементы ИОС");

            return Enumerable.Empty<CheckCommandError>();
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            #region Подготовливаю спец. классы
            int totalLength = elemColl.Length;
            // Инженерные элементы разделяю на 2 части и 2 таски. Часть 1
            CheckMEPHeightMEPData[] mepDataColl1 = null;
            Task prepareMEPDataTask1 = Task.Run(() =>
            {
                mepDataColl1 = new ArraySegment<Element>(elemColl, 0, totalLength/2)
                    .Select(e => new CheckMEPHeightMEPData(e).SetCurrentSolidColl().SetCurrentBBoxColl())
                    // Проверка элементов на предмет наличия геометрии
                    .Where(m => m.MEPElemSolids.Count != 0)
                    .ToArray();
            });
            
            // Часть 2
            CheckMEPHeightMEPData[] mepDataColl2 = null;
            Task prepareMEPDataTask2 = Task.Run(() =>
            {
                mepDataColl2 = new ArraySegment<Element>(elemColl, totalLength / 2, totalLength - totalLength / 2)
                    .Select(e => new CheckMEPHeightMEPData(e).SetCurrentSolidColl().SetCurrentBBoxColl())
                    // Проверка элементов на предмет наличия геометрии
                    .Where(m => m.MEPElemSolids.Count != 0)
                    .ToArray();
            });

            // Архитектурные элементы помещений
            CheckMEPHeightARRoomData[] checkMEPHeightARData = null;
            Task checkMEPHeightARDataTask = Task.Run(() =>
            {
                checkMEPHeightARData = PrepareARRoomData(doc);
            });

            Task.WaitAll(prepareMEPDataTask1, prepareMEPDataTask2, checkMEPHeightARDataTask);
            
            CheckMEPHeightMEPData[] mepDataColl = mepDataColl1.Concat(mepDataColl2).ToArray();
            #endregion

            #region Обработка элементов ИОС
            // Анализ элементов ИОС на элементы АР
            foreach (CheckMEPHeightARRoomData arRoomData in checkMEPHeightARData)
            {
                #region Инженерные элементы с привязкой к помещениям - разделяю на 2 части и 2 таски
                totalLength = mepDataColl.Length;
                CheckMEPHeightMEPData[] currentRoomMEPDataColl1 = null;
                Task currentRoomMEPDataCollTask1 = Task.Run(() =>
                {
                    currentRoomMEPDataColl1 = new ArraySegment<CheckMEPHeightMEPData>(mepDataColl, 0, totalLength / 2)
                        .Where(mep => mep.IsElemInCurrentRoom(arRoomData))
                        .ToArray();
                });

                CheckMEPHeightMEPData[] currentRoomMEPDataColl2 = null;
                Task currentRoomMEPDataCollTask2 = Task.Run(() =>
                {
                    currentRoomMEPDataColl2 = new ArraySegment<CheckMEPHeightMEPData>(mepDataColl, totalLength / 2, totalLength - totalLength / 2)
                        .Where(mep => mep.IsElemInCurrentRoom(arRoomData))
                        .ToArray();
                });

                Task.WaitAll(currentRoomMEPDataCollTask1, currentRoomMEPDataCollTask2);
                CheckMEPHeightMEPData[] currentRoomMEPDataColl = currentRoomMEPDataColl1.Concat(currentRoomMEPDataColl2).ToArray();
                #endregion

                #region Инженерные элементы с нарушением высоты - разделяю на 2 части и 2 таски
                totalLength = currentRoomMEPDataColl.Length;
                CheckMEPHeightMEPData[] errorMEPDataColl1 = null;
                Task errorMEPDataCollTask1 = Task.Run(() =>
                {
                    errorMEPDataColl1 = CheckMEPHeightMEPData
                        .CheckIOSElemsForMinDistErrorByAR(new ArraySegment<CheckMEPHeightMEPData>(currentRoomMEPDataColl, 0, totalLength / 2), arRoomData);
                });

                CheckMEPHeightMEPData[] errorMEPDataColl2 = null;
                Task errorMEPDataCollTask2 = Task.Run(() =>
                {
                    errorMEPDataColl2 = CheckMEPHeightMEPData
                        .CheckIOSElemsForMinDistErrorByAR(new ArraySegment<CheckMEPHeightMEPData>(currentRoomMEPDataColl, totalLength / 2, totalLength - totalLength / 2), arRoomData);
                });

                Task.WaitAll(errorMEPDataCollTask1, errorMEPDataCollTask2);
                CheckMEPHeightMEPData[] errorMEPDataColl = errorMEPDataColl1.Concat(errorMEPDataColl2).ToArray();
                #endregion

                List<Element> verticalCurveElemsFiltered_ErrorElemsColl = new List<Element>();
                foreach (CheckMEPHeightMEPData mepData in errorMEPDataColl)
                {
                    bool isVerticalElem = false;
                    if (mepData.MEPElement is InsulationLiningBase insLining)
                    {
                        isVerticalElem = CheckMEPHeightMEPData.VerticalCurveElementsFilteredWithTolerance(doc.GetElement(insLining.HostElementId), errorMEPDataColl);
                        if (isVerticalElem)
                            verticalCurveElemsFiltered_ErrorElemsColl.Add(mepData.MEPElement);
                    }
                    else
                    {
                        isVerticalElem = CheckMEPHeightMEPData.VerticalCurveElementsFilteredWithTolerance(mepData.MEPElement, errorMEPDataColl);
                        if (isVerticalElem)
                            verticalCurveElemsFiltered_ErrorElemsColl.Add(mepData.MEPElement);
                    }
                }

                if (verticalCurveElemsFiltered_ErrorElemsColl.Any())
                {
                    result.Add(new WPFEntity(
                        verticalCurveElemsFiltered_ErrorElemsColl,
                        CheckStatus.Error,
                        $"Недопустимая дистанция для помещения {arRoomData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()}: {arRoomData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()}",
                        $"Минимально допустимая высота монтажа элементов по версии ГИ: {Math.Round((arRoomData.RoomMinDistance * 304.8), 0)}",
                        true,
                        false));
                }
            }
            #endregion

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Метод подготовки элементов АР для последующего анализа
        /// </summary>
        /// <param name="doc"></param>
        /// <exception cref="CheckerException"></exception>
        private CheckMEPHeightARRoomData[] PrepareARRoomData(Document doc)
        {
            //Подготовка связей АР
            RevitLinkInstance[] arLinkInsts = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                // Слабое место - имена файлов могут отличаться из-за требований Заказчика
                .Where(lm =>
                    (lm.Name.ToUpper().Contains("_AR_") || lm.Name.ToUpper().Contains("_АР_"))
                    || (lm.Name.ToUpper().Contains("_AR.RVT") || lm.Name.ToUpper().Contains("_АР.RVT"))
                    || (lm.Name.ToUpper().StartsWith("AR_") || lm.Name.ToUpper().StartsWith("АР_")))
                .Cast<RevitLinkInstance>()
                .ToArray();

            // Проверка на то, чтобы файлы АР подверглись поиску по паттерну из проверки выше
            if (arLinkInsts.Count() == 0)
                throw new CheckerException("Не удалось идентифицировать связи - они либо названы не по внутреннему BEP KPLN (обр. в BIM-отдел), либо связи в модели отсутсвуют (подгрузи)");

            // Проверка на то, чтобы ВСЕ файлы АР были открыты в модели
            if (arLinkInsts.Where(rli => rli.GetLinkDocument() == null).Any())
                throw new CheckerException("Перед запуском - открой все связи АР");

            // Получаю список назначений помещений, которые необходимо проверить (ПОВЕСИТЬ НА КОНФИГ!!!!!)
            string[] roomDepartmentColl =
            {
                "Кладовая",
                "МОП",
                "Технические помещения"
            };

            // Получаю список части имен помещений, которые НЕ являются ошибками (ПОВЕСИТЬ НА КОНФИГ!!!!!)
            string[] roomNameExceptionColl =
            {
                "итп",
                "пространство",
                "насосн",
                "камера",
            };

            return CheckMEPHeightARRoomData.PreapareMEPHeightARRoomDataColl(arLinkInsts, roomDepartmentColl, roomNameExceptionColl);
        }


        /// <summary>
        /// Получение элементов ИОС по списку категорий, с учетом фильтрации
        /// </summary>
        private Element[] PreapareIOSElements(Document doc)
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
            List<FilteredElementCollector> bicColl = new List<FilteredElementCollector>(_mepBICColl.Count);
            foreach (BuiltInCategory bic in _mepBICColl)
            {
                FilteredElementCollector fic = new FilteredElementCollector(doc).OfCategory(bic);
                fic.WherePasses(eFilter).WhereElementIsNotElementType();
                bicColl.Add(fic);
            }

            // Добавляю очищенные элементы в коллекцию
            foreach (FilteredElementCollector coll in bicColl)
            {
                result.AddRange(coll.ToElements());
            }

            return result.ToArray();
        }
    }
}
