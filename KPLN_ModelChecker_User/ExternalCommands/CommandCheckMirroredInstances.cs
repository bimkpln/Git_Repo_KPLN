using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckMirroredInstances : AbstrCheckCommand<CommandCheckMirroredInstances>, IExternalCommand
    {
        internal const string PluginName = "Проверка зеркальных эл-в";

        /// <summary>
        /// Список категорий для проверки
        /// </summary>
        private readonly List<BuiltInCategory> _bicErrorSearch = new List<BuiltInCategory>()
        {
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_CurtainWallPanels,
            BuiltInCategory.OST_MechanicalEquipment
        };

        /// <summary>
        /// Список исключений в именах, при проверке
        /// </summary>
        private readonly List<string> _exceptionMEPFamilyNameList = new List<string>()
        {
            "556_",
            "557_"
        };

        public CommandCheckMirroredInstances() : base()
        {
        }

        internal CommandCheckMirroredInstances(ExtensibleStorageEntity esEntity) : base(esEntity)
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
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            List<Element> checkElemColl = new List<Element>();
            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType();

                // У оборудования нужно брать только элементы из списка
                if (bic == BuiltInCategory.OST_MechanicalEquipment) checkElemColl.AddRange(FilteredByStringContainsColl(bicColl).ToElements());
                // У остального - берем все, кроме семейств проемов
                else checkElemColl.AddRange(
                    bicColl
                    .Cast<Element>()
                    .Where(e => !(doc.GetElement(e.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsElementId()) as ElementType).FamilyName.StartsWith("100_Проем")));
            }

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, checkElemColl.ToArray());
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

            foreach (Element element in elemColl)
            {
                if (!(element is FamilyInstance instance))
                {
                    // Стены могут выступать в качестве панелей витража. Зеркальность тут проверять не нужно
                    if (element is Wall wall) continue;
                    else throw new Exception($"У элемента с id: {element.Id} - невозможно взять FamilyInstance.");
                }

                // Для панелей витража анализируются ТОЛЬКО окна и двери. Также для них нужно брать основание - host.
                // Основание дополнительно проверятся на поворот - flip
                if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_CurtainWallPanels)
                {
                    string elName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString();
                    if (elName.StartsWith("135_") && elName.ToLower().Contains("двер")
                        || (elName.ToLower().Contains("створк") && !elName.ToLower().Contains("глух")))
                    {
                        Wall panelHostWall = instance.Host as Wall;
                        if (panelHostWall.Flipped)
                        {
                            WPFEntity hostEntity = new WPFEntity(
                                ESEntity,
                                element,
                                "Недопустимый зеркальный элемент",
                                "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                                string.Empty,
                                true,
                                SetApproveStatusByUserComment(element, CheckStatus.Error),
                                true);

                            result.Add(hostEntity);
                        }
                    }
                }
                else
                {
                    if (instance.Mirrored && instance.SuperComponent == null)
                    {
                        WPFEntity elemEntity = new WPFEntity(
                            ESEntity,
                            element,
                            "Недопустимый зеркальный элемент",
                            "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                            string.Empty,
                            true,
                            SetApproveStatusByUserComment(element, CheckStatus.Error),
                            true);

                        result.Add(elemEntity);
                    }
                }
            }

            return result
                .OrderBy(e =>
                    ((Level)doc.GetElement(e.Element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId())).Elevation)
                .ToList();
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByCategory();
        }

        /// <summary>
        /// Метод для создания фильтра, для игнорирования элементов по имени семейства
        /// </summary>
        private FilteredElementCollector FilteredByStringContainsColl(FilteredElementCollector currentColl)
        {
            List<ElementFilter> resFilterColl = new List<ElementFilter>();
            foreach (string currentName in _exceptionMEPFamilyNameList)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateContainsRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, false);
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                resFilterColl.Add(eFilter);
            }
            
            currentColl.WherePasses(new LogicalOrFilter(resFilterColl));
            return currentColl;
        }
    }
}
