using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckMirroredInstances : AbstrCheck
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
        /// Список имен семейств ИОС для проверки
        /// </summary>
        private readonly string[] _mepFamilyNames = new string[9]
        {
            // Для стандартных шаблонов КПЛН
            "556_",
            "557_",
            
            // Для МТРС (делали НЕ на нашем шаблоне)
            "KZTO_Гармония_Радиатор",
            "Радиатор панельный SPL CC",
            "Радиатор панельный SPL CV",
            "Внутрипольный конвектор с вентилятором SPL Instyle FC Standart",

            // Для СЕТ (делали НЕ на нашем шаблоне)
            "ASML_ОВ_Конвектор_Внутрипольный_Бриз_КЗТО",
            "ASML_ОВ_Радиатор_Универсальный_Боковое подключение",
            "ASML_ОВ_Радиатор_Универсальный_Нижнее подключение",
        };

        public CheckMirroredInstances() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка зеркальных эл-в";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckMirroredInstances",
                    new Guid("33b660af-95b8-4d7c-ac42-c9425320557b"),
                    new Guid("33b660af-95b8-4d7c-ac42-c9425320557c"));
        }

        public override Element[] GetElemsToCheck()
        {
            List<Element> result = new List<Element>();

            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(CheckDocument)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType();

                // У оборудования нужно брать только элементы из списка
                if (bic == BuiltInCategory.OST_MechanicalEquipment)
                    result.AddRange(FilteredByString_BeginsWith(bicColl).ToElements());
                // У остального - берем все, кроме семейств проемов
                else result.AddRange(
                    bicColl
                    .Cast<Element>()
                    .Where(e =>
                        !(CheckDocument.GetElement(e.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsElementId()) as ElementType).FamilyName.StartsWith("100_Проем")));
            }

            return result.ToArray();
        }

        private protected override IEnumerable<CheckerEntity> GetCheckerEntities(Element[] elemColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (Element element in elemColl)
            {
                if (!(element is FamilyInstance instance))
                {
                    // Стены могут выступать в качестве панелей витража. Зеркальность тут проверять не нужно
                    if (element is Wall wall)
                        continue;
                    else
                        throw new Exception($"У элемента с id: {element.Id} - невозможно взять FamilyInstance.");
                }

                // Для панелей витража анализируются ТОЛЬКО окна и двери. Также для них нужно брать основание - host.
                // Основание дополнительно проверятся на поворот - flip
                if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_CurtainWallPanels)
                {
                    string elName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString();
                    if (elName.StartsWith("135_")
                        && elName.ToLower().Contains("двер")
                        || (elName.ToLower().Contains("створк") && !elName.ToLower().Contains("глух")))
                    {
                        Wall panelHostWall = instance.Host as Wall;
                        if (panelHostWall.Flipped)
                        {
                            CheckerEntity hostEntity = new CheckerEntity(
                                element,
                                "Недопустимый зеркальный элемент",
                                "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                                string.Empty,
                                true)
                                .Set_CanApproved()
                                .Set_DataByESData(ESEntity);


                            result.Add(hostEntity);
                        }
                    }
                }
                else
                {
                    if (instance.Mirrored && instance.SuperComponent == null)
                    {
                        CheckerEntity elemEntity = new CheckerEntity(
                            element,
                            "Недопустимый зеркальный элемент",
                            "Указанный элемент запрещено зеркалить, т.к. это повлияет на выдаваемые объемы в спецификациях",
                            string.Empty,
                            true)
                            .Set_CanApproved()
                            .Set_DataByESData(ESEntity);

                        result.Add(elemEntity);
                    }
                }
            }

            return result.OrderBy(e =>
                    ((Level)CheckDocument.GetElement(e.Element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId())).Elevation);
        }

        /// <summary>
        /// Метод для создания фильтра, для игнорирования элементов по имени семейства (НАЧИНАЕТСЯ С)
        /// </summary>
        private FilteredElementCollector FilteredByString_BeginsWith(FilteredElementCollector currentColl)
        {
            List<ElementFilter> resFilterColl = new List<ElementFilter>();
            foreach (string currentName in _mepFamilyNames)
            {
#if Debug2020 || Revit2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, false);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName);
#endif
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                resFilterColl.Add(eFilter);
            }

            currentColl.WherePasses(new LogicalOrFilter(resFilterColl));
            return currentColl;
        }
    }
}
