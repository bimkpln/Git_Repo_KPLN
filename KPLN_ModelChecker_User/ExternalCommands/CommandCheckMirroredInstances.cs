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
    internal class CommandCheckMirroredInstances : AbstrCheckCommand, IExternalCommand
    {
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

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            _name = "Проверка зеркальных элементов";
            _application = uiapp;

            _allStorageName = "KPLN_CheckMirroredInstances";

            _lastRunGuid = new Guid("33b660af-95b8-4d7c-ac42-c9425320447b");
            _userTextGuid = new Guid("33b660af-95b8-4d7c-ac42-c9425320447c");

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

        private protected override List<CheckCommandError> CheckElements(Document doc, Element[] elemColl) => null;

        private protected override List<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
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
                            BoundingBoxXYZ bbox = instance.get_BoundingBox(null);
                            // У панелей витража при моделировании возникают фантомы. Они не идут в спеку, их можно не анализировать
                            if (bbox == null) continue;

                            WPFEntity hostEntity = new WPFEntity(
                                element,
                                SetApproveStatusByUserComment(element, Status.Error),
                                "Недопустимый зеркальный элемент",
                                "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                                true,
                                true);

                            hostEntity.PrepareZoomGeometryExtension(bbox);
                            result.Add(hostEntity);
                        }
                    }
                }
                else
                {
                    if (instance.Mirrored && instance.SuperComponent == null)
                    {
                        WPFEntity elemEntity = new WPFEntity(
                            element,
                            SetApproveStatusByUserComment(element, Status.Error),
                            "Недопустимый зеркальный элемент",
                            "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                            true,
                            true);
                        elemEntity.PrepareZoomGeometryExtension(instance.get_BoundingBox(null));
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
            foreach (string currentName in _exceptionMEPFamilyNameList)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateContainsRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                currentColl.WherePasses(eFilter);
            }
            return currentColl;
        }
    }
}
