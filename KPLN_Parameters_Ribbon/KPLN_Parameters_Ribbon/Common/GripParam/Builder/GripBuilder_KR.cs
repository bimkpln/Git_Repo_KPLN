﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_KR : AbstrGripBuilder
    {
        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        public override bool Prepare()
        {
            // Категория "Стены" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => !x.Name.ToLower().Contains("перепад") || !x.Name.ToLower().Contains("балк")));

            // Категория "Стены" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => x.Name.ToLower().Contains("перепад") || x.Name.ToLower().Contains("балк")));

            // Категория "Перекрытия" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => !x.Name.ToLower().Contains("площадка") && !x.Name.ToLower().Contains("фундамент") && !x.Name.ToLower().Contains("пандус")));

            // Категория "Перекрытия" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => x.Name.ToLower().Contains("площадка") || x.Name.ToLower().Contains("фундамент") || x.Name.ToLower().Contains("пандус")));

            // Семейства "Обобщенные модели" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => !i.Symbol.FamilyName.StartsWith("22") && i.Symbol.FamilyName.StartsWith("2")));

            // Семейства "Окна" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => i.Symbol.FamilyName.StartsWith("23")));

            // Категория "Лестницы" над уровнем и в отдельный список
            IEnumerable<StairsRun> stairsRun = new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsRun))
               .Cast<StairsRun>();
            IEnumerable<StairsLanding> stairsLanding = new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsLanding))
               .Cast<StairsLanding>();
            ElemsOnLevel.AddRange(stairsRun);
            ElemsOnLevel.AddRange(stairsLanding);
            StairsElems.AddRange(stairsRun);
            StairsElems.AddRange(stairsLanding);

            // Семейства "Колоны" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>());
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Cast<FamilyInstance>());

            // Семейства "Колоны" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>());


            // Семейства "Перекрытия" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Floors)
                .Cast<FamilyInstance>());

            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<FamilyInstance>());

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

            FloorNumberUnderLevelByElement(pb);

            return true;
        }

        public override bool ExecuteSectionParams(Progress_Single pb)
        {
            bool byElem = false;
            bool byUnderElem = false;
            bool byStairsElem = false;

            if (ElemsOnLevel.Count > 0)
            {
                byElem = SectionExcecuter.ExecuteByElement(Doc, ElemsOnLevel, "Орг.ОсьБлок", SectionParamName, pb);
            }

            if (ElemsByHost.Count > 0)
            {
                byUnderElem = SectionExcecuter.ExecuteByElement(Doc, ElemsUnderLevel, "Орг.ОсьБлок", SectionParamName, pb);
            }

            if (StairsElems.Count > 0)
            {
                byStairsElem = SectionExcecuter.ExecuteByElement(Doc, StairsElems, "Орг.ОсьБлок", SectionParamName, pb);
            }

            return byElem && byUnderElem && byStairsElem;
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
        /// Заполняю номер этажа для элементов, находящихся ПОД уровнем
        /// </summary>
        protected virtual void FloorNumberUnderLevelByElement(Progress_Single pb)
        {
            foreach (Element elem in ElemsUnderLevel)
            {
                Level baseLevel = LevelTool.GetLevelOfElement(elem, Doc);
                if (baseLevel != null)
                {
                    string floorNumber = null;

                    double offsetFromLev = LevelTool.GetElementLevelGrip(elem, baseLevel);

                    if (offsetFromLev < 0)
                    {
                        floorNumber = LevelTool.GetFloorNumberByLevel(baseLevel, LevelNumberIndex, SplitLevelChar);
                    }
                    else
                    {
                        floorNumber = LevelTool.GetFloorNumberIncrementLevel(baseLevel, LevelNumberIndex, Doc, SplitLevelChar);
                    }

                    if (floorNumber == null)
                    {
                        Print($"Не найден уровень выше, для уровня {baseLevel.Name} " +
                            $"при обработке элемента: {elem.Name} c id: {elem.Id.IntegerValue}." +
                            "\nДля уровней необходимо заполнить параметр: На уровень выше, за исключением последнего этажа",
                            KPLN_Loader.Preferences.MessageType.Error);

                        continue;
                    }
                    Parameter floor = elem.LookupParameter(LevelParamName);
                    if (floor == null) continue;
                    floor.Set(floorNumber);
                    
                    pb.Increment();
                }
                else
                {
                    Print($"Не найден уровень у элемента с Id: {elem.Id}", KPLN_Loader.Preferences.MessageType.Error);
                }
            }
        }
    }
}
