using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckMEPHeight : AbstrCheckCommand, IExternalCommand
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
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            _name = "Проверка высоты эл-в ИОС";
            _application = uiapp;

            _allStorageName = "KPLN_CheckMEPHeight";

            _lastRunGuid = new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b68");
            _userTextGuid = new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b69");

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            Element[] mepELems = PreapareIOSElements(doc);

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, mepELems);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override List<CheckCommandError> CheckElements(Document doc, Element[] elemColl)
        {
            if (!elemColl.Any())
                throw new UserException("В проекте отсутсвуют необходимые элементы ИОС");

            return null;
        }

        private protected override List<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            #region Подготовка элементов АР
            //Подготовка связей АР

            List<RevitLinkInstance> arLinkInsts = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                // Слабое место - имена файлов могут отличаться из-за требований Заказчика
                .Where(lm =>
                    (lm.Name.ToUpper().Contains("_AR_") || lm.Name.ToUpper().Contains("_АР_"))
                    || (lm.Name.ToUpper().Contains("_AR.RVT") || lm.Name.ToUpper().Contains("_АР.RVT"))
                    || (lm.Name.ToUpper().StartsWith("AR_") || lm.Name.ToUpper().StartsWith("АР_")))
                .Cast<RevitLinkInstance>()
                .ToList();

            // Проверка на то, чтобы ВСЕ файлы АР были открыты в модели
            if (arLinkInsts.Where(rli => rli.GetLinkDocument() == null).Any())
                throw new UserException("Перед запуском - открой все связи АР");

            List<CheckMEPHeightARRoomData> checkMEPHeightARData = CheckMEPHeightARRoomData.PreapareMEPHeightARRoomDataColl(arLinkInsts);
            return null;
            #endregion

            #region Обработка элементов ИОС
            // Подготовка элементов ИОС
            CheckMEPHeightMEPData[] mepDataColl = elemColl
                .Select(e => new CheckMEPHeightMEPData(e))
                .Where(m => m.CurrentSolid != null)
                .ToArray();

            // Анализ элементов ИОС на элементы АР
            int cnt = 0;
            foreach (CheckMEPHeightARRoomData arData in checkMEPHeightARData)
            {
                cnt++;

                CheckMEPHeightMEPData[] currentRoomMEPDataColl = mepDataColl
                    .Where(mep => CheckMEPHeightMEPData.IsElemInCurrentRoomCheck(mep, arData))
                    .ToArray();

                CheckMEPHeightMEPData[] errorMEPDataColl = CheckMEPHeightMEPData.CheckIOSElemsForMinDistErrorByAR(currentRoomMEPDataColl, arData);

                List<Element> verticalCurveElemsFiltered_ErrorElemsColl = new List<Element>();
                foreach (CheckMEPHeightMEPData mepData in errorMEPDataColl)
                {
                    bool isVerticalElem = false;
                    if (mepData.CurrentElement is InsulationLiningBase insLining)
                    {
                        isVerticalElem = CheckMEPHeightMEPData.VerticalCurveElementsFilteredWithTolerance(doc.GetElement(insLining.HostElementId), errorMEPDataColl);
                        if (isVerticalElem)
                            verticalCurveElemsFiltered_ErrorElemsColl.Add(mepData.CurrentElement);
                    }
                    else
                    {

                        isVerticalElem = CheckMEPHeightMEPData.VerticalCurveElementsFilteredWithTolerance(mepData.CurrentElement, errorMEPDataColl);
                        if (isVerticalElem)
                            verticalCurveElemsFiltered_ErrorElemsColl.Add(mepData.CurrentElement);
                    }
                }

                if (verticalCurveElemsFiltered_ErrorElemsColl.Any())
                {
                    result.Add(new WPFEntity(
                        verticalCurveElemsFiltered_ErrorElemsColl,
                        Status.Error,
                        $"Недопустимая дистанция для помещения {arData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()}: {arData.CurrentRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()}",
                        $"Минимально допустимая высота монтажа элементов по версии ГИ: {Math.Round((arData.RoomMinDistance * 304.8), 0)}",
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
