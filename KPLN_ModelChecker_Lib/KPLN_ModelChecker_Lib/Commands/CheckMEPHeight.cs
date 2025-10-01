using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.CheckMEPHeight;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_Lib.Forms;
using KPLN_ModelChecker_Lib.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckMEPHeight : AbstrCheck
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

        /// <summary>
        /// Коллекция сущностей для окна параметров помещений (далее будет синхрон со спец классами эл-тов моделей - CheckMEPHeightARRoomData)
        /// </summary>
        private CMHEntity[] _cmhEntites;
        private CheckMEPHeightARRoomData[] _checkMEPHeightARData;

        public CheckMEPHeight() : base()
        {
            if (PluginName == null)
                PluginName = "ИОС: Проверка высоты эл-в";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckMEPHeight",
                    new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b68"),
                    new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b69"));
        }

        public override Element[] GetElemsToCheck()
        {
            List<Element> result = new List<Element>();

            // Генерация фильтров
            List<FilterRule> filtRules = new List<FilterRule>(_exceptionFamilyStartNameList.Count);
            foreach (string currentName in _exceptionFamilyStartNameList)
            {
#if Debug2020 || Revit2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName);
#endif
                filtRules.Add(fRule);
            }
            ElementParameterFilter eFilter = new ElementParameterFilter(filtRules);

            // Генерация и фильтрация FilteredElementCollector
            List<FilteredElementCollector> bicColl = new List<FilteredElementCollector>(_mepBICColl.Count);
            foreach (BuiltInCategory bic in _mepBICColl)
            {
                FilteredElementCollector fic = new FilteredElementCollector(CheckDocument).OfCategory(bic);
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

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            #region Подготовка связей АР
            RevitLinkInstance[] arLinkInsts = new FilteredElementCollector(CheckDocument)
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
            #endregion

            _cmhEntites = CheckMEPHeightARRoomData.PreapareMEPHeightViewModelDataDataColl(arLinkInsts);
            CheckMEPHeight_ARDataForm arDataFrom = new CheckMEPHeight_ARDataForm(CheckDocument, _cmhEntites);
            if (!(bool)arDataFrom.ShowDialog())
                return CheckResultStatus.Cancelled;

            var onlyCheckedVMColl = arDataFrom.ViewModelColl.Where(v => v.VMIsCheckRun);
            _checkMEPHeightARData = CheckMEPHeightARRoomData.PreapareMEPHeightARRoomDataColl(onlyCheckedVMColl, arLinkInsts);


            #region Подготовливаю спец. классы
            int totalLength = elemColl.Length;
            // Инженерные элементы разделяю на 2 части и 2 таски. Часть 1
            CheckMEPHeightMEPData[] mepDataColl1 = null;
            Task prepareMEPDataTask1 = Task.Run(() =>
            {
                mepDataColl1 = new ArraySegment<Element>(elemColl, 0, totalLength / 2)
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

            Task.WaitAll(prepareMEPDataTask1, prepareMEPDataTask2);

            CheckMEPHeightMEPData[] mepDataColl = mepDataColl1.Concat(mepDataColl2).ToArray();
            #endregion

            #region Обработка элементов ИОС
            // Анализ элементов ИОС на элементы АР
            foreach (CheckMEPHeightARRoomData arRoomData in _checkMEPHeightARData)
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
                        isVerticalElem = CheckMEPHeightMEPData.VerticalCurveElementsFilteredWithTolerance(CheckDocument.GetElement(insLining.HostElementId), errorMEPDataColl);
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
                    _checkerEntitiesCollHeap.Add(new CheckerEntity(
                        verticalCurveElemsFiltered_ErrorElemsColl,
                        $"Недопустимая дистанция для помещения {arRoomData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()}: {arRoomData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()}",
                        $"Минимально допустимая высота монтажа элементов для данного помещения: {Math.Round((arRoomData.CurrentRoomMinDistance), 0)}",
                        string.Empty,
                        false));
                }
            }
            #endregion

            return CheckResultStatus.Succeeded;
        }
    }
}