using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_AR : AbstrGripBuilder
    {
        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        public override bool Prepare()
        {
            // Категория "Стены"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => !x.Name.StartsWith("00_")));

            // Категория "Перекрытия"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => !x.Name.StartsWith("00_")));

            // Категория "Кровля"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(RoofBase))
                .Cast<RoofBase>());

            // Семейства "Окна"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilyInstance>());

            // Семейства "Двери"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>());

            // Семейства "Панели витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                .Cast<FamilyInstance>());

            // Семейства "Импосты витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallMullions)
                .Cast<FamilyInstance>());

            // Семейства "Лестничные марши"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Railing))
                .Cast<Railing>());

            // Семейства "Оборудование"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                .Where(x => !x.Name.StartsWith("199_")));

            // Семейства "Обощенные модели"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => !x.Name.StartsWith("199_")));

            // Семейства "Каркас несущий (перемычки)"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>()
                .Where(x => !x.Name.StartsWith("199_")));


            AllElementsCount = ElemsOnLevel.Count + ElemsUnderLevel.Count + StairsElems.Count;

            if (AllElementsCount > 0)
            {
                return true;
            }
            else
            {
                throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет, или имя проекта не соответсвует ВЕР!");
            }
        }

        public override bool ExecuteLevelParams(Progress_Single pb)
        {
            LevelTool.Levels = new FilteredElementCollector(Doc)
                .OfClass(typeof(Level))
                .WhereElementIsNotElementType()
                .Cast<Level>();

            FloorNumberOnLevelByElement(pb);

            FloorNumberOnLevelByHost(pb);

            return true;
        }

        public override bool ExecuteSectionParams(Progress_Single pb)
        {
            bool byElem = false;
            bool byHost = false;
            bool byStairsElem = false;

            if (ElemsOnLevel.Count > 0)
            {
                byElem = SectionExcecuter.ExecuteByElement(Doc, ElemsOnLevel, "КП_О_Секция", SectionParamName, pb);
            }

            if (ElemsByHost.Count > 0)
            {
                byHost = SectionExcecuter.ExecuteByHost_AR(ElemsByHost, SectionParamName, pb);
            }

            if (StairsElems.Count > 0)
            {
                byStairsElem = SectionExcecuter.ExecuteByElement(Doc, StairsElems, "КП_О_Секция", SectionParamName, pb);
            }

            return byElem && byHost && byStairsElem;
        }

        /// <summary>
        /// Заполняю номер этажа для элементов, находящихся НА уровне
        /// </summary>
        protected virtual void FloorNumberOnLevelByElement(Progress_Single pb)
        {
            foreach (Element elem in ElemsOnLevel)
            {
                try
                {
                    Level baseLevel = LevelTool.GetLevelOfElement(elem, Doc);
                    if (baseLevel != null)
                    {
                        string floorNumber = LevelTool.GetFloorNumberByLevel(baseLevel, LevelNumberIndex, SplitLevelChar);

                        double offsetFromLev = LevelTool.GetElementLevelGrip(elem, baseLevel);

                        if (offsetFromLev < 0)
                        {
                            floorNumber = LevelTool.GetFloorNumberDecrementLevel(baseLevel, LevelNumberIndex, SplitLevelChar);
                        }

                        if (floorNumber == null) continue;
                        Parameter floor = elem.LookupParameter(LevelParamName);
                        if (floor == null) continue;
                        floor.Set(floorNumber);
                        
                        pb.Increment();
                    }
                }
                catch (Exception e)
                {
                    PrintError(e, "Не удалось обработать элемент: " + elem.Id.IntegerValue + " " + elem.Name);
                }
            }
        }

        /// <summary>
        /// Заполняю номер этажа для вложенных элементов
        /// </summary>
        protected virtual void FloorNumberOnLevelByHost(Progress_Single pb)
        {
            foreach (Element elem in ElemsByHost)
            {
                Element hostElem = null;
                Wall hostWall = null;
                Type elemType = elem.GetType();
                switch (elemType.Name)
                {
                    // Проброс панелей витража на стену
                    case nameof(Panel):
                        Panel panel = (Panel)elem;
                        hostWall = (Wall)panel.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс импостов витража на стену
                    case nameof(Mullion):
                        Mullion mullion = (Mullion)elem;
                        hostWall = (Wall)mullion.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс на вложенные общие семейства
                    default:
                        FamilyInstance instance = elem as FamilyInstance;
                        hostElem = instance.SuperComponent;
                        break;
                }

                var hostElemParamValue = hostElem.LookupParameter(LevelParamName).AsString();
                elem.LookupParameter(LevelParamName).Set(hostElemParamValue);

                pb.Increment();
            }
        }
    }
}
