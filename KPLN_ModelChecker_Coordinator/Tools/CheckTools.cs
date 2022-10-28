extern alias revit;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_Coordinator.Common.Collections;
using revit.Autodesk.Revit.DB;
using System.IO;
using KPLN_ModelChecker_Coordinator.Common;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Coordinator.Tools
{
    public static class CheckTools
    {
        public static BuiltInCategory[] CategoriesToCheck = new BuiltInCategory[] {
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_GenericModel };

        public static int CheckFileSize(string path)
        {
            FileInfo fileInfo = new FileInfo(path);
            return (int)Math.Round((double)fileInfo.Length / 1048576);
        }

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private static string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            "max",
            "min"
        };

        public static int CheckErrors(Document doc)
        {
            List<string> false_warnings_list = new List<string>() {
                "В одном и том же месте имеются идентичные экземпляры.",
                "Включите выноску или перенесите марку",
                "Выделенные этажи перекрывают друг друга.",
                "Выделенные стены перекрываются.",
                "Выделенные элементы объединены, но они не пересекаются.",
                "Стена слегка отклонилась от оси и может вызвать неточности.",
                "Величина толщины пола может быть неточной из-за избыточного редактирования формы.",
                "Элементы имеют повторяющиеся значения «Номер».",
                "Для одной или нескольких балясин верхний опорный объект находится ниже нижнего. Эти балясины не были созданы.",
                "Ограждения полностью размещены вне основы. Можно скорректировать ограждение, передвинув его на основу и изменив вертикальную позицию.",
                "Линии-разделители выделенных помещений перекрываются",
                "Перекрытие объемов Помещение",
                "Рассчитанная фактическая высота высота подступенка больше максимальной высоты подступенка, заданной для данного типа лестницы",
                "Рассчитанная фактическая ширина проступи меньше минимальной ширины проступи, заданной для данного типа лестницы",
                "Глубина площадки меньше ширины марша."
            };
            
            int counter = 0;
            
            IList<FailureMessage> warnings = doc.GetWarnings();
            {
                foreach (FailureMessage fmsg in warnings)
                {
                    string fmsg_str = fmsg.GetDescriptionText();
                    foreach (string fwl in false_warnings_list)
                    {
                        if (fmsg_str.Contains(fwl))
                        {
                            counter++;
                        }
                    }

                }
            }
            return counter;
        }
        
        public static int CheckLevels(Document doc)
        {
            double bpOffset = 0;
            foreach (BasePoint bp in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).ToElements())
            {
                bpOffset = bp.get_BoundingBox(null).Min.Z;
            }
            int ammount = 0;
            try
            {
                string code = null;
                switch (LevelChecker.CheckLevels(doc))
                {
                    case CheckResult.NoSections:
                        code = null;
                        break;
                    case CheckResult.Error:
                        return new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements().Count;
                    case CheckResult.Corpus:
                        code = "К";
                        break;
                    case CheckResult.Sections:
                        code = "С";
                        break;
                }
                LevelChecker.Levels.Clear();
                
                foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
                {
                    LevelChecker.AddLevel(element as Level, doc, code);
                }
                
                foreach (BuiltInCategory cat in CategoriesToCheck)
                {
                    foreach (Element element in new FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements())
                    {
                        if (!element.ElementPassesConditions()) continue;

                        if (element.GetType() == typeof(FamilyInstance))
                        {
                            // Игнорирование вложенных экземпляров в семейство
                            if ((element as FamilyInstance).SuperComponent != null) continue;
                        }

                        try
                        {
                            CalculateType linkType = CalculateType.Default;
                            Level level = doc.GetElement(element.LevelId) as Level;

                            #region Category and level
                            if (level == null)
                            {
                                try
                                {
                                    level = doc.GetElement(element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).AsElementId()) as Level;
                                }
                                catch (Exception) { }
                            }

                            if (element.GetType() == typeof(FamilyInstance))
                            {
                                FamilyInstance familyInstance = (FamilyInstance)element;
                                if (familyInstance != null)
                                {
                                    Element host = familyInstance.Host;
                                    if (host != null)
                                    {
                                        if (familyInstance.Host.GetType() == typeof(Floor))
                                        {
                                            linkType = CalculateType.Floor;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (element.GetType() == typeof(Floor)) linkType = CalculateType.Floor;
                            }

                            if (level == null) { continue; }

                            #endregion
                            BoundingBoxXYZ box = element.get_BoundingBox(null);
                            // Игнорирование элементов, у которых нет геометрии (например - панели витражей)
                            if (box == null)
                            {
                                var a = element.Id.IntegerValue;
                                continue;
                            }

                            BoundingBoxXYZ boxAnalitical = new BoundingBoxXYZ() { Min = box.Min - new XYZ(0, 0, bpOffset), Max = box.Max - new XYZ(0, 0, bpOffset) };

                            LevelChecker checker = LevelChecker.GetLevelById(level.Id);
                            LevelCheckResult result = linkType == CalculateType.Default ? checker.GetLevelIntersection(boxAnalitical) : checker.GetFloorLevelIntersection(boxAnalitical);
                            switch (result)
                            {
                                case LevelCheckResult.NotInside:
                                    if (Check(linkType, LevelCheckResult.FullyInside, level, boxAnalitical, doc))
                                    {
                                        ammount++;
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception) { }
            return ammount;
        }

        public static bool Check(CalculateType linkType, LevelCheckResult result, Level level, BoundingBoxXYZ box, Document doc)
        {
            foreach (LevelChecker c in LevelChecker.GetOtherLevelById(level.Id))
            {
                LevelCheckResult rslt = linkType == CalculateType.Default ? c.GetLevelIntersection(box) : c.GetFloorLevelIntersection(box);

                if (rslt == LevelCheckResult.FullyInside)
                {
                    // Игнорирую для КР привязку элементов к уровню выше (так выдаются спеки)
                    if (doc.Title.ToLower().Contains("_кр_") || doc.Title.ToLower().Contains("_kr_") || doc.Title.ToLower().Contains("_kg_"))
                    {
                        if (level.Id == c.Level.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId())
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            return false;
        }

        public static int CheckMirrored(Document doc)
        {
            int ammount = 0;
            if (doc.Title.Contains("KR") || doc.Title.Contains("КР"))
            {
                return ammount;
            }
            try
            {
                foreach (BuiltInCategory category in new BuiltInCategory[] { BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows, BuiltInCategory.OST_CurtainWallPanels })
                {
                    foreach (Element element in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(category).WhereElementIsNotElementType().ToElements())
                    {
                        try
                        {
                            FamilyInstance instance = element as FamilyInstance;
                            BuiltInCategory enumCat = (BuiltInCategory)element.Category.Id.IntegerValue;
                            if (enumCat == BuiltInCategory.OST_CurtainWallPanels)
                            {
                                string elName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                                if (elName.StartsWith("135_") && elName.ToLower().Contains("двер") | elName.ToLower().Contains("створк"))
                                {
                                    Wall panelHostWall = instance.Host as Wall;
                                    if (panelHostWall.Flipped)
                                    {
                                        ammount++;
                                    }
                                }
                            }
                            else
                            {
                                if (instance.Mirrored)
                                {
                                    ammount++;
                                }
                            }
                        }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception e)
            { PrintError(e); }
            return ammount;
        }
        
        public static int CheckSharedLocations(Document doc)
        {
            int ammount = 0;
            try
            {
                ///HashSet<string> links = new HashSet<string>();
                foreach (RevitLinkInstance link in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements())
                {
                    string[] separators = { ".rvt : " };
                    string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                    int lenNS = nameSubs.Length;
                    if (lenNS > 2)
                    {
                        continue;
                    }
                    try
                    {
                        Document linkDocument = link.GetLinkDocument();
                        string name = linkDocument.PathName;
                        string currentPosition = link.Name.Split(new string[] { "позиция " }, StringSplitOptions.RemoveEmptyEntries).Last();
                        /*
                        Это была проверка на дубликаты экземпляров связей. Ревит не позволяет нескольким экземплярам связей иметь одну и ту же площадку. Т.е. такая связь попадет в ошибку с отсутствием площадки
                        if (!links.Contains(name))
                        {
                            links.Add(name);
                        }
                        else
                        {
                            ammount++;
                        }
                        */
                        if (currentPosition == "<Не общедоступное>")
                        {
                            ammount++;
                        }
                        else
                        {
                            bool detected = false;
                            foreach (ProjectLocation i in doc.ProjectLocations)
                            {
                                if (i.Name == currentPosition)
                                {
                                    detected = true;
                                    if (currentPosition == "Встроенный")
                                    {
                                        ammount++;
                                    }
                                    else
                                    {
                                        if (!link.Pinned)
                                        {
                                            ammount++;

                                        }
                                    }
                                }
                            }
                            if (!detected)
                            {
                                foreach (ProjectLocation i in linkDocument.ProjectLocations)
                                {
                                    if (i.Name == currentPosition)
                                    {
                                        detected = true;
                                        if (currentPosition == "Встроенный")
                                        {
                                            ammount++;
                                        }
                                        else
                                        {
                                            if (!link.Pinned)
                                            {
                                                ammount++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            ammount++;
                        }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception e)
            { PrintError(e); }
            return ammount;
        }
        
        public static List<List<Element>> GetElementsCollection(Document doc)
        {
            List<Element> families = new List<Element>();
            List<Element> symbols = new List<Element>();
            foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).ToElements())
            {
                if (family.IsEditable && family.IsUserCreated)
                {
                    families.Add(family);
                    foreach (ElementId symbolId in family.GetFamilySymbolIds())
                    {
                        FamilySymbol symbol = doc.GetElement(symbolId) as FamilySymbol;
                        symbols.Add(symbol);
                    }
                }
            }
            return new List<List<Element>> { families, symbols };
        }

        
        public static bool AllWorksetsAreOpened(Document doc)
        {
            foreach (Workset w in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
            {
                if (!w.IsOpen)
                {
                    return false;
                }
            }
            return true;
        }
        
        public static int CheckLinkWorkSets(Document doc)
        {
            int ammount = 0;
            try
            {
                List<Workset> worksets = new List<Workset>();
                foreach (Workset w in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
                {
                    if (!w.IsOpen)
                    {
                        ammount++;
                    }
                    else
                    {
                        worksets.Add(w);
                    }
                }
                foreach (RevitLinkInstance link in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements())
                {
                    string[] separators = { ".rvt : " };
                    string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                    int lenNS = nameSubs.Length;
                    if (lenNS > 3)
                    {
                        continue;
                    }
                    foreach (Workset w in worksets)
                    {
                        if (link.WorksetId.IntegerValue == w.Id.IntegerValue)
                        {
                            if (!w.Name.StartsWith("00") & !w.Name.StartsWith("#"))
                            {
                                ammount++;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            { PrintError(e); }
            return ammount;
        }
        
        public static int CheckElementsWorksets(Document doc)
        {
            int ammount = 0;
            try
            {
                HashSet<string> links = new HashSet<string>();
                if (doc.IsWorkshared)
                {
                    List<Workset> worksets = new List<Workset>();
                    foreach (Workset w in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
                    {
                        if (!w.IsOpen)
                        {
                            ammount++;
                            return ammount;
                        }
                        worksets.Add(w);
                    }
                    foreach (Element element in new FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())
                    {
                        if (element.Category == null) { continue; }
                        
                        try
                        {
                            if (element.GetType() == typeof(RevitLinkInstance) || element.GetType() == typeof(ImportInstance))
                            {
                                continue;
                            }

                            if ((element.Category.CategoryType == CategoryType.Annotation) & (element.GetType() == typeof(Grid) | element.GetType() == typeof(Level)))
                            {
                                string wsName = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                                if (!wsName.ToLower().Contains("оси и уровни") & !wsName.ToLower().Contains("общие уровни и сетки"))
                                {
                                    ammount++;
                                }
                            }

                            // Есть внутренняя ошибка Revit, когда появляются компоненты легенды, которые нигде не размещены, и у них редактируемый рабочий набор. Вручную такой элемент - создать НЕВОЗМОЖНО
                            if (element.Category.CategoryType == CategoryType.Model && element.Category.Id.IntegerValue != -2000576)
                            {
                                foreach (Workset w in worksets)
                                {
                                    if (element.WorksetId.IntegerValue == w.Id.IntegerValue)
                                    {
                                        // Проверка замонитренных моделируемых элементов
                                        if (element.GetMonitoredLinkElementIds().Count() > 0)
                                        {
                                            if (!w.Name.StartsWith("02"))
                                            {
                                                ammount++;

                                                continue;
                                            }
                                        }

                                        // Проверка остальных моделируемых элементов на рабочий набор связей 
                                        else if (w.Name.StartsWith("00")
                                            | w.Name.StartsWith("#")
                                            && !w.Name.Contains("DWG"))
                                        {
                                            ammount++;

                                            continue;
                                        }

                                        // Проверка остальных моделируемых элементов на рабочий набор для сеток
                                        else if (w.Name.ToLower().Contains("оси и уровни")
                                            | w.Name.ToLower().Contains("общие уровни и сетки"))
                                        {
                                            ammount++;

                                            continue;
                                        }

                                        // Проверка остальных моделируемых элементов на рабочий набор для связей
                                        else if (w.Name.StartsWith("02"))
                                        {
                                            ammount++;

                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }
                else
                {
                    ammount++;
                    return ammount;
                }
            }
            catch (Exception)
            { }
            return ammount;
        }

        public static int CheckMonitorGrids(Document doc)
        {
            int ammount = 0;
            try
            {
                HashSet<int> ids = new HashSet<int>();
                foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements())
                {
                    try
                    {
                        if (element.IsMonitoringLinkElement())
                        {
                            RevitLinkInstance link = null;
                            List<string> names = new List<string>();
                            foreach (ElementId i in element.GetMonitoredLinkElementIds())
                            {
                                ids.Add(i.IntegerValue);
                                link = doc.GetElement(i) as RevitLinkInstance;
                                names.Add(link.Name);
                            }
                            if (link == null)
                            {
                                ammount++;
                            }
                        }
                        else
                        {
                            ammount++;
                        }
                    }
                    catch (Exception)
                    { }
                }
                if (ids.Count > 1)
                {
                    ammount++;
                }
            }
            catch (Exception e)
            { PrintError(e); }
            return ammount;
        }
        
        public static int CheckMonitorLevels(Document doc)
        {
            int ammount = 0;
            try
            {
                HashSet<int> ids = new HashSet<int>();
                foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
                {
                    try
                    {
                        if (element.IsMonitoringLinkElement())
                        {
                            RevitLinkInstance link = null;
                            List<string> names = new List<string>();
                            foreach (ElementId i in element.GetMonitoredLinkElementIds())
                            {
                                ids.Add(i.IntegerValue);
                                link = doc.GetElement(i) as RevitLinkInstance;
                                names.Add(link.Name);
                            }
                            if (link == null)
                            {
                                ammount++;
                            }
                        }
                        else
                        {
                            ammount++;
                        }
                    }
                    catch (Exception)
                    { }
                }
                if (ids.Count > 1)
                {
                    ammount++;
                }
            }
            catch (Exception e)
            { PrintError(e); }
            return ammount;
        }
        
        public static int CheckFamilies(Document doc)
        {
            int ammount = 0;

            CheckFamiliesNames(doc, ref ammount);

            FilteredElementCollector docDimensions = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).WhereElementIsNotElementType();
            CheckOverrideDimension(doc, docDimensions, ref ammount);

            CheckAccuracyDemension(doc, ref ammount);

            return ammount;
        }

        /// <summary>
        /// Поиск похожего имени. Одинаковым должна быть только первичная часть имени, до среза по циферным значениям
        /// </summary>
        /// <param name="currentName">Имя, которое нужно проанализировать</param>
        /// <param name="elemsColl">Коллекция, по которой нужно осуществлять поиск</param>
        /// <returns>Имя подобного элемента</returns>
        private static string SearchSimilarName(string currentName, List<Element> elemsColl)
        {
            string similarFamilyName = String.Empty;

            // Осуществляю поиск цифр в конце имени
            string digitEndTrimmer = Regex.Match(currentName, @"\d*$").Value;
            // Осуществляю срез имени на найденные цифры в конце имени
            string truePartOfName = currentName.TrimEnd(digitEndTrimmer.ToArray());
            if (digitEndTrimmer.Length > 0)
            {
                foreach (Element checkElem in elemsColl)
                {
                    if (!checkElem.Equals(currentName) && checkElem.Name.Equals(truePartOfName.TrimEnd(new char[] { ' ' })))
                    {
                        similarFamilyName = checkElem.Name;
                    }
                }
            }
            return similarFamilyName;
        }

        /// <summary>
        /// Проверка имен семейств и типоразмеров
        /// </summary>
        private static void CheckFamiliesNames(Document doc, ref int ammount)
        {
            try
            {
                List<Element> docFamilies = new List<Element>();
                HashSet<int> familyIds = new HashSet<int>();
                HashSet<string> docFamilyNames = new HashSet<string>();

                foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).ToElements())
                {
                    if (!familyIds.Contains(family.Id.IntegerValue))
                    {
                        docFamilies.Add(family);
                        familyIds.Add(family.Id.IntegerValue);
                        docFamilyNames.Add(family.Name);
                    }
                    else
                    {
                        continue;
                    }
                }

                foreach (Family currentFam in docFamilies)
                {
                    List<Element> currentFamilySymols = new List<Element>();
                    string currentFamName = currentFam.Name;
                    if (Regex.Match(currentFamName, @"\b[.0]\d*$").Value.Length > 2)
                    {
                        ammount++;
                    }

                    string similarFamilyName = SearchSimilarName(currentFamName, docFamilies);
                    if (!similarFamilyName.Equals(String.Empty))
                    {
                        ammount++;
                    }

                    foreach (ElementId id in currentFam.GetFamilySymbolIds())
                    {
                        FamilySymbol symbol = currentFam.Document.GetElement(id) as FamilySymbol;
                        currentFamilySymols.Add(symbol);
                    }

                    foreach (FamilySymbol currentSymbol in currentFamilySymols)
                    {
                        string currentSymName = currentSymbol.Name;
                        string similarSymbolName = SearchSimilarName(currentSymName, currentFamilySymols);

                        if (!similarSymbolName.Equals(String.Empty))
                        {
                            ammount++;
                        }
                    }
                }
            }
            catch (Exception)
            { }
        }

        /// <summary>
        /// Проверка размеров
        /// </summary>
        private static void CheckOverrideDimension(Document doc, FilteredElementCollector docDimensions, ref int ammount)
        {
            foreach (Dimension dim in docDimensions)
            {
                // Игнорирую чертежные виды
                try
                {
                    if (dim.View.GetType().Equals(typeof(ViewDrafting))) { continue; }
                }
                catch (NullReferenceException) { continue; }

                double? currentValue = dim.Value;

                if (currentValue.HasValue && dim.ValueOverride?.Length > 0)
                {
                    double value = currentValue.Value * 304.8;
                    string overrideValue = dim.ValueOverride;
                    int dimId = dim.Id.IntegerValue;
                    string dimName = dim.Name;
                    CheckDimValues(doc, value, overrideValue, dimId, dimName, ref ammount);
                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        if (segment.ValueOverride?.Length > 0)
                        {
                            double value = segment.Value.Value * 304.8;
                            string overrideValue = segment.ValueOverride;
                            int dimId = dim.Id.IntegerValue;
                            string dimName = dim.Name;
                            CheckDimValues(doc, value, overrideValue, dimId, dimName, ref ammount);
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Метод для анализа значений размеров
        /// </summary>
        private static void CheckDimValues(Document doc, double value, string overrideValue, int elemId, string elemName, ref int ammount)
        {

            string ovverrideMinValue = String.Empty;
            double overrideMinDouble = 0.0;
            string ovverrideMaxValue = String.Empty;
            double overrideMaxDouble = 0.0;

            string[] splitValues = overrideValue.Split(_separArr, StringSplitOptions.None);

            // Анализирую диапозоны
            if (splitValues.Length > 1)
            {
                ovverrideMinValue = splitValues[0];
                if (ovverrideMinValue.Length == 0)
                {
                    ovverrideMinValue = overrideValue;
                }
                else
                {
                    ovverrideMaxValue = splitValues[1];
                }

                string onlyNumbMin = new string(ovverrideMinValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideMinDouble);
                if (!ovverrideMaxValue.Equals(String.Empty))
                {
                    string onlyNumbMax = new string(ovverrideMaxValue.Where(x => Char.IsDigit(x)).ToArray());
                    Double.TryParse(onlyNumbMax, out overrideMaxDouble);
                    // Нахожу значения вне диапозоне
                    if (value >= overrideMaxDouble | value < overrideMinDouble)
                    {
                        ammount++;
                    }
                }
            }

            // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
            else
            {
                double overrideDouble = 0.0;
                string onlyNumbMin = new string(overrideValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideDouble);
                if (overrideDouble == 0.0)
                {
                    ammount++;
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    ammount++;
                }
            }
        }

        /// <summary>
        /// Метод для анализа округлений размеров
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ammount"></param>
        private static void CheckAccuracyDemension(Document doc, ref int ammount)
        {
            string docTitle = doc.Title;
            FilteredElementCollector docDimensionTypes = new FilteredElementCollector(doc).OfClass(typeof(DimensionType)).WhereElementIsElementType();
            foreach (DimensionType dimType in docDimensionTypes)
            {
                if (dimType.UnitType == UnitType.UT_Length)
                {
                    FormatOptions typeOpt = dimType.GetUnitsFormatOptions();
                    try
                    {
                        double currentAccuracy = typeOpt.Accuracy;
                        if (currentAccuracy > 1.0)
                        {
                            ammount++;
                        }
                    }
                    catch (Exception)
                    {
                        //Игнорирую типы без настроек округления
                    }
                }
            }
        }
    }
}
