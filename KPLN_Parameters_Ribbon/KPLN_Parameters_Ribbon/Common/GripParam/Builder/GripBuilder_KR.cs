using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Parameters_Ribbon.Common.GripParam.Actions;
using KPLN_Parameters_Ribbon.Common.Tools;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_KR : AbstrGripBuilder
    {

        /// <summary>
        /// Элементы под уровнем
        /// </summary>
        protected List<Element> _elemsUnderLevel = new List<Element>();

        /// <summary>
        /// Коллекция всех лестниц
        /// </summary>
        protected List<Element> _stairsElems = new List<Element>();

        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        public override bool Prepare()
        {
            // Категория "Стены" над уровнем
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => !x.Name.ToLower().Contains("перепад") || !x.Name.ToLower().Contains("балк")));

            // Категория "Стены" под уровнем
            _elemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => x.Name.ToLower().Contains("перепад") || x.Name.ToLower().Contains("балк")));

            // Категория "Перекрытия" над уровнем
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => !x.Name.ToLower().Contains("площадка") && !x.Name.ToLower().Contains("фундамент") && !x.Name.ToLower().Contains("пандус")));

            // Категория "Перекрытия" под уровнем
            _elemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => x.Name.StartsWith("00_"))
                .Where(x => x.Name.ToLower().Contains("площадка") || x.Name.ToLower().Contains("фундамент") || x.Name.ToLower().Contains("пандус")));

            // Семейства "Обобщенные модели" над уровнем
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => !i.Symbol.FamilyName.StartsWith("22") && i.Symbol.FamilyName.StartsWith("2")));

            // Семейства "Окна" над уровнем
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
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
            _elemsOnLevel.AddRange(stairsRun);
            _elemsOnLevel.AddRange(stairsLanding);
            _stairsElems.AddRange(stairsRun);
            _stairsElems.AddRange(stairsLanding);

            // Семейства "Колоны" над уровнем
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>());
            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Cast<FamilyInstance>());

            // Семейства "Колоны" под уровнем
            _elemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>());


            // Семейства "Перекрытия" под уровнем
            _elemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Floors)
                .Cast<FamilyInstance>());

            _elemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<FamilyInstance>());

            _allElems.AddRange(_elemsOnLevel);
            _allElems.AddRange(_elemsUnderLevel);
            _allElems.AddRange(_stairsElems);

            if (_allElems.Count > 0)
            {
                return true;
            }
            else
            {
                throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет!");
            }
        }

        public override bool ExecuteLevelParams()
        {
            int counter = 0;

            FloorNumberOnLevel(ref counter);

            FloorNumberUnderLevel(ref counter);

            Print("Обработано элементов: " + counter, KPLN_Loader.Preferences.MessageType.Warning);

            return true;
        }

        public override bool ExecuteSectionParams()
        {
            return SectionExcecuter.Execute(Doc, _allElems, "Орг.ОсьБлок", SectionParamName);
        }

        /// <summary>
        /// Заполняю номер этажа для элементов, находящихся НА уровне
        /// </summary>
        protected virtual void FloorNumberOnLevel(ref int counter)
        {
            foreach (Element elem in _elemsOnLevel)
            {
                try
                {
                    Level baseLevel = LevelTool.GetLevelOfElement(elem, Doc);
                    if (baseLevel != null)
                    {
                        string floorNumber = LevelTool.GetFloorNumberByLevel(baseLevel, 1, SplitLevelChar);
                        if (floorNumber == null) continue;
                        Parameter floor = elem.LookupParameter(LevelParamName);
                        if (floor == null) continue;
                        floor.Set(floorNumber);
                        counter++;
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
        protected virtual void FloorNumberUnderLevel(ref int counter)
        {
            foreach (Element elem in _elemsUnderLevel)
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
                    counter++;
                }
                else
                {
                    Print($"Не найден уровень у элемента с Id: {elem.Id}", KPLN_Loader.Preferences.MessageType.Error);
                }
            }
        }
    }
}
