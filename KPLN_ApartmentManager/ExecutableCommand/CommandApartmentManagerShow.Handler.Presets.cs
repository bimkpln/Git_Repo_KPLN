using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ApartmentManager.Common;
using KPLN_ApartmentManager.ExecutableCommand;
using KPLN_ApartmentManager.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private const int ShaftWallTypeStorageKey = int.MinValue;
        private const int LoggiaWallTypeStorageKey = int.MinValue + 1;

        private void ExecuteRefreshApartmentPresets(Document doc, ApartmentPresetData currentPreset)
        {
            ViewPlan activeFloorPlan = doc.ActiveView as ViewPlan;
            if (activeFloorPlan == null || activeFloorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "Перед обновлением данных откройте план этажа.");
                return;
            }

            ApartmentPresetPanelContext context = BuildPresetPanelContext(doc, activeFloorPlan);

            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                _window.SetApartmentPresetContext(
                    context,
                    currentPreset != null ? currentPreset.Clone() : null);
            }));
        }

        private ApartmentPresetPanelContext BuildPresetPanelContext(Document doc, ViewPlan activeFloorPlan)
        {
            ApartmentPresetPanelContext context = new ApartmentPresetPanelContext();
            context.ActivePlanName = activeFloorPlan != null ? activeFloorPlan.Name : "";

            List<ViewPlan> plans = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(x => !x.IsTemplate && x.ViewType == ViewType.FloorPlan && x.GenLevel != null)
                .OrderBy(x => activeFloorPlan == null || x.Id != activeFloorPlan.Id)
                .ThenBy(x => x.Name)
                .ToList();

            foreach (ViewPlan plan in plans)
            {
                if (activeFloorPlan != null && plan.Id == activeFloorPlan.Id)
                {
                    context.Plans.Add(BuildResolvedPlanPresetOption(doc, plan.Name));
                    continue;
                }

                context.Plans.Add(new ApartmentPlanPresetOption
                {
                    PlanName = plan.Name,
                    IsResolved = false
                });
            }

            context.ResolvePlanData = delegate (string planName)
            {
                return BuildResolvedPlanPresetOption(doc, planName);
            };

            return context;
        }

        private ApartmentPlanPresetOption BuildResolvedPlanPresetOption(Document doc, string planName)
        {
            ApartmentPlanPresetOption option = new ApartmentPlanPresetOption();
            option.PlanName = planName;
            option.IsResolved = true;

            ViewPlan plan = FindTargetFloorPlan(doc, planName);
            if (plan == null)
            {
                option.LowerConstraintText = "";
                option.UpperConstraintText = "Неприсоединённая";
                option.RoomCategories = new List<string> { "Помещение" };
                option.WindowTypeOptions = BuildWindowTypeOptions(doc);
                option.HasWindowMarkers = false;
                option.ShaftWallTypeOptions = BuildAllWallTypeOptions(doc);
                option.HasShaftWallMarkers = false;
                option.LoggiaWallTypeOptions = BuildAllWallTypeOptions(doc);
                option.HasLoggiaWallMarkers = false;
                return option;
            }

            option.LowerConstraintText = BuildLowerConstraintTextForPlan(doc, plan);
            option.UpperConstraintText = "Неприсоединённая";
            option.ModelSignature = BuildApartmentPlanModelSignature(doc, plan);
            option.WallThicknesses = BuildWallThicknessesForPlan(doc, plan);
            option.WallTypeOptionsByThickness = BuildWallTypeOptionsByThicknessForPlan(doc, plan);
            option.WindowTypeOptions = BuildWindowTypeOptions(doc);
            option.HasWindowMarkers = HasWindowMarkersForPlan(doc, plan);
            option.ShaftWallTypeOptions = BuildAllWallTypeOptions(doc);
            option.HasShaftWallMarkers = HasShaftWallMarkersForPlan(doc, plan);
            option.LoggiaWallTypeOptions = BuildAllWallTypeOptions(doc);
            option.HasLoggiaWallMarkers = HasLoggiaWallMarkersForPlan(doc, plan);
            option.RoomCategories = BuildRoomCategoriesForPlan(doc, plan);

            if (option.RoomCategories == null || option.RoomCategories.Count == 0)
                option.RoomCategories = new List<string> { "Помещение" };

            option.DoorRequirements = BuildDoorRequirementsForPlan(doc, plan);
            option.DoorTypeOptionsByRequirementKey = BuildDoorTypeOptionsByRequirementForPlan(doc, option.DoorRequirements);

            return option;
        }

        private static string BuildApartmentPlanModelSignature(Document doc, ViewPlan plan)
        {
            if (doc == null || plan == null)
                return "";

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan)
                .Where(x => x != null)
                .OrderBy(x => IDHelper.ElIdValue(x.Id))
                .ToList();

            if (apartments.Count == 0)
                return "count=0";

            return "count=" + apartments.Count + ";" + string.Join(
                ";",
                apartments.Select(x =>
                    IDHelper.ElIdValue(x.Id) +
                    ":loggia=" +
                    (IsLoggiaEnabled(x) ? "1" : "0")));
        }

        private string BuildLowerConstraintTextForPlan(Document doc, ViewPlan plan)
        {
            if (plan == null)
                return "Нет ни одного 2D-семейства";

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return "Нет ни одного 2D-семейства";

            HashSet<string> levelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance fi in apartments)
            {
                ElementId levelId = GetInstanceLevelId(fi);
                if (levelId == ElementId.InvalidElementId)
                    continue;

                Level lvl = doc.GetElement(levelId) as Level;
                if (lvl != null && !string.IsNullOrWhiteSpace(lvl.Name))
                    levelNames.Add(lvl.Name);
            }

            if (levelNames.Count == 0)
                return "Нет ни одного 2D-семейства";

            return string.Join(", ", levelNames.OrderBy(x => x));
        }

        private static ViewPlan FindTargetFloorPlan(Document doc, string selectedPlanName)
        {
            List<ViewPlan> plans = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(x => !x.IsTemplate && x.ViewType == ViewType.FloorPlan && x.GenLevel != null)
                .ToList();

            if (!string.IsNullOrWhiteSpace(selectedPlanName))
            {
                ViewPlan selected = plans.FirstOrDefault(x =>
                    string.Equals(x.Name, selectedPlanName, StringComparison.OrdinalIgnoreCase));

                if (selected != null)
                    return selected;
            }

            ViewPlan activePlan = doc.ActiveView as ViewPlan;
            if (activePlan != null && !activePlan.IsTemplate && activePlan.ViewType == ViewType.FloorPlan && activePlan.GenLevel != null)
                return activePlan;

            return null;
        }

        private List<string> BuildRoomCategoriesForPlan(Document doc, ViewPlan plan)
        {
            HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (plan == null || plan.GenLevel == null)
                return new List<string> { "Помещение" };

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesOnLevel(doc, plan.GenLevel);

            if (apartments.Count == 0)
                return new List<string> { "Помещение" };

            foreach (FamilyInstance apartment in apartments)
            {
                List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartment);

                foreach (FamilyInstance roomFi in rooms)
                {
                    string categoryName = GetRoomCategoryLabel(roomFi);
                    if (!string.IsNullOrWhiteSpace(categoryName))
                        result.Add(categoryName);
                }
            }

            if (result.Count == 0)
                return new List<string> { "Помещение" };

            return result.OrderBy(x => x).ToList();
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesForPlan(Document doc, ViewPlan plan)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            if (doc == null || plan == null)
                return result;

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (FamilyInstance fi in instances)
            {
                if (fi == null)
                    continue;

                Parameter pComment = fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pComment == null)
                    continue;

                string comment = pComment.AsString();
                if (string.IsNullOrWhiteSpace(comment))
                    continue;

                if (!comment.Contains(ApartmentInstanceMarker))
                    continue;

                bool belongsToPlan = false;

                if (fi.OwnerViewId != ElementId.InvalidElementId && fi.OwnerViewId == plan.Id)
                    belongsToPlan = true;

                if (!belongsToPlan && plan.GenLevel != null)
                {
                    ElementId levelId = GetInstanceLevelId(fi);
                    if (levelId != ElementId.InvalidElementId && levelId == plan.GenLevel.Id)
                        belongsToPlan = true;
                }

                if (belongsToPlan)
                    result.Add(fi);
            }

            return result;
        }

        private static List<FamilyInstance> FindRoomSubComponents(Document doc, FamilyInstance apartmentInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (apartmentInstance == null)
                return result;

            ICollection<ElementId> subIds = apartmentInstance.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return result;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                string familyName = "";
                string typeName = "";

                if (subFi.Symbol != null)
                {
                    typeName = subFi.Symbol.Name ?? "";
                    if (subFi.Symbol.Family != null)
                        familyName = subFi.Symbol.Family.Name ?? "";
                }

                Category cat = subFi.Category;

                bool byFamily = familyName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;
                bool byType = typeName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;
                bool byCategory = cat != null && cat.Name.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;

                if (byFamily || byType || byCategory)
                    result.Add(subFi);
            }

            return result;
        }

        private static string GetRoomCategoryLabel(FamilyInstance roomFi)
        {
            if (roomFi == null)
                return null;

            string[] candidateParams = new[]
            {
                "Категория помещения",
                "Категория",
                "Назначение",
                "Имя",
                "Наименование"
            };

            foreach (string paramName in candidateParams)
            {
                Parameter p = roomFi.LookupParameter(paramName);
                if (p != null)
                {
                    string value = p.AsString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }
            }

            if (roomFi.Symbol != null && !string.IsNullOrWhiteSpace(roomFi.Symbol.Name))
                return roomFi.Symbol.Name.Trim();

            if (roomFi.Symbol != null && roomFi.Symbol.Family != null && !string.IsNullOrWhiteSpace(roomFi.Symbol.Family.Name))
                return roomFi.Symbol.Family.Name.Trim();

            if (roomFi.Category != null && !string.IsNullOrWhiteSpace(roomFi.Category.Name))
                return roomFi.Category.Name.Trim();

            return "Помещение";
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesOnLevel(Document doc, Level level)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (FamilyInstance fi in instances)
            {
                Parameter pComment = fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pComment == null)
                    continue;

                string comment = pComment.AsString();
                if (string.IsNullOrWhiteSpace(comment))
                    continue;

                if (!comment.Contains(ApartmentInstanceMarker))
                    continue;

                ElementId levelId = GetInstanceLevelId(fi);
                if (levelId == ElementId.InvalidElementId)
                    continue;

                if (levelId == level.Id)
                    result.Add(fi);
            }

            return result;
        }

        private List<int> BuildWallThicknessesForPlan(Document doc, ViewPlan plan)
        {
            HashSet<int> result = new HashSet<int>();

            if (plan == null)
                return new List<int>();

            List<FamilyInstance> apartmentsOnPlan = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartmentsOnPlan.Count == 0)
                return new List<int>();

            foreach (FamilyInstance fi in apartmentsOnPlan)
            {
                double thicknessInternal;
                if (!TryGetApartmentWallThickness(fi, out thicknessInternal))
                    continue;

                int thicknessMm = (int)Math.Round(IDHelper.ConvertInternalToMm(thicknessInternal));
                if (thicknessMm > 0)
                    result.Add(thicknessMm);
            }

            return result.OrderBy(x => x).ToList();
        }

        private static bool TryGetApartmentWallThickness(FamilyInstance apartmentFi, out double thicknessInternal)
        {
            thicknessInternal = 0;

            if (apartmentFi == null)
                return false;

            Parameter p = apartmentFi.LookupParameter("Стены_Толщина");
            if (p == null)
                p = apartmentFi.LookupParameter("Стены Толщина");

            if (p != null && p.StorageType == StorageType.Double)
            {
                thicknessInternal = p.AsDouble();
                return true;
            }

            Element typeElem = apartmentFi.Document.GetElement(apartmentFi.GetTypeId());
            if (typeElem != null)
            {
                p = typeElem.LookupParameter("Стены_Толщина");
                if (p == null)
                    p = typeElem.LookupParameter("Стены Толщина");

                if (p != null && p.StorageType == StorageType.Double)
                {
                    thicknessInternal = p.AsDouble();
                    return true;
                }
            }

            return false;
        }

        private static double GetApartmentWallThickness(FamilyInstance apartmentFi)
        {
            if (apartmentFi == null)
                throw new ArgumentNullException("apartmentFi");

            Parameter p = apartmentFi.LookupParameter("Стены_Толщина");
            if (p == null)
                p = apartmentFi.LookupParameter("Стены Толщина");

            if (p != null && p.StorageType == StorageType.Double)
                return p.AsDouble();

            Element typeElem = apartmentFi.Document.GetElement(apartmentFi.GetTypeId());
            if (typeElem != null)
            {
                p = typeElem.LookupParameter("Стены_Толщина");
                if (p == null)
                    p = typeElem.LookupParameter("Стены Толщина");

                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            throw new Exception("Не найден параметр 'Стены_Толщина' у экземпляра или типа семейства квартиры.");
        }

        private Dictionary<int, List<string>> BuildWallTypeOptionsByThicknessForPlan(Document doc, ViewPlan plan)
        {
            Dictionary<int, List<string>> result = new Dictionary<int, List<string>>();

            List<int> thicknesses = BuildWallThicknessesForPlan(doc, plan);
            if (thicknesses.Count == 0)
                return result;

            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            foreach (int thicknessMm in thicknesses)
            {
                List<string> matchedNames = new List<string>();

                foreach (WallType wt in allWallTypes)
                {
                    if (wt == null || string.IsNullOrWhiteSpace(wt.Name))
                        continue;

                    int wallTypeThicknessMm;
                    if (!TryGetWallTypeThicknessMm(wt, out wallTypeThicknessMm))
                        continue;

                    if (wallTypeThicknessMm == thicknessMm)
                        matchedNames.Add(wt.Name);
                }

                result[thicknessMm] = matchedNames
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();
            }

            return result;
        }

        private List<string> BuildAllWallTypeOptions(Document doc)
        {
            if (doc == null)
                return new List<string>();

            return new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                .Select(x => x.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        private static string GetPresetShaftWallType(ApartmentPresetData preset)
        {
            if (preset == null)
                return "Не выбрано";

            string value;
            if (preset.WallTypeByThickness != null &&
                preset.WallTypeByThickness.TryGetValue(ShaftWallTypeStorageKey, out value) &&
                !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            try
            {
                var prop = preset.GetType().GetProperty("ShaftWallType");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    value = prop.GetValue(preset, null) as string;
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }
            catch
            {
            }

            return "Не выбрано";
        }

        private static string GetPresetLoggiaWallType(ApartmentPresetData preset)
        {
            if (preset == null)
                return "Не выбрано";

            string value;
            if (preset.WallTypeByThickness != null &&
                preset.WallTypeByThickness.TryGetValue(LoggiaWallTypeStorageKey, out value) &&
                !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            try
            {
                var prop = preset.GetType().GetProperty("LoggiaWallType");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    value = prop.GetValue(preset, null) as string;
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }
            catch
            {
            }

            return "Не выбрано";
        }

        private static bool TryGetWallTypeThicknessMm(WallType wallType, out int thicknessMm)
        {
            thicknessMm = 0;

            if (wallType == null)
                return false;

            Parameter p = wallType.LookupParameter("Толщина");
            if (p != null && p.StorageType == StorageType.Double)
            {
                thicknessMm = (int)Math.Round(IDHelper.ConvertInternalToMm(p.AsDouble()));
                return thicknessMm > 0;
            }

            try
            {
                double width = wallType.Width;
                if (width > 0)
                {
                    thicknessMm = (int)Math.Round(IDHelper.ConvertInternalToMm(width));
                    return thicknessMm > 0;
                }
            }
            catch
            {
            }

            return false;
        }

        private List<string> BuildWindowTypeOptions(Document doc)
        {
            if (doc == null)
                return new List<string>();

            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Select(BuildWindowTypeDisplayName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        private bool HasWindowMarkersForPlan(Document doc, ViewPlan plan)
        {
            if (doc == null || plan == null)
                return false;

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return false;

            foreach (FamilyInstance apartmentFi in apartments)
            {
                List<FamilyWindowMarker> windows = CollectWindowMarkersFromApartmentInstance(doc, apartmentFi);
                if (windows.Count > 0)
                    return true;
            }

            return false;
        }

        private bool HasShaftWallMarkersForPlan(Document doc, ViewPlan plan)
        {
            if (doc == null || plan == null)
                return false;

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return false;

            foreach (FamilyInstance apartmentFi in apartments)
            {
                List<FamilyShaftWallMarker> shafts = CollectShaftWallMarkersFromApartmentInstance(doc, apartmentFi);
                if (shafts.Count > 0)
                    return true;
            }

            return false;
        }

        private bool HasLoggiaWallMarkersForPlan(Document doc, ViewPlan plan)
        {
            if (doc == null || plan == null)
                return false;

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return false;

            foreach (FamilyInstance apartmentFi in apartments)
            {
                List<FamilyLoggiaWallMarker> loggias = CollectLoggiaWallMarkersFromApartmentInstance(doc, apartmentFi);
                if (loggias.Count > 0)
                    return true;
            }

            return false;
        }

        private List<ApartmentDoorRequirementOption> BuildDoorRequirementsForPlan(Document doc, ViewPlan plan)
        {
            Dictionary<string, ApartmentDoorRequirementOption> result =
                new Dictionary<string, ApartmentDoorRequirementOption>(StringComparer.OrdinalIgnoreCase);

            if (doc == null || plan == null)
                return new List<ApartmentDoorRequirementOption>();

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return new List<ApartmentDoorRequirementOption>();

            foreach (FamilyInstance apartmentFi in apartments)
            {
                if (apartmentFi == null)
                    continue;

                List<FamilyInstance> doorInstances = FindDoorSubComponentsRecursive(doc, apartmentFi);

                foreach (FamilyInstance doorFi in doorInstances)
                {
                    if (doorFi == null)
                        continue;

                    string typeName = doorFi.Symbol != null ? doorFi.Symbol.Name ?? "" : "";
                    string commentValue = GetCommentsValue(doorFi);
                    bool isEntranceDoor = HasEntranceDoorComment(doorFi);

                    int widthMm;
                    if (!TryGetDoorWidthMmFrom2DMarker(doorFi, typeName, out widthMm) || widthMm <= 0)
                        continue;

                    AddDoorRequirement(result, typeName, commentValue, widthMm, isEntranceDoor);
                }

            }

            return result.Values
                .OrderBy(x => !x.IsEntranceDoor)
                .ThenBy(x => x.RoomCategory)
                .ThenBy(x => x.WidthMm)
                .ThenBy(x => x.DoorTypeName2D)
                .ToList();
        }

        private static void AddDoorRequirement(Dictionary<string, ApartmentDoorRequirementOption> result, string typeName, string commentValue,
            int widthMm, bool isEntranceDoor)
        {
            if (result == null || string.IsNullOrWhiteSpace(typeName) || widthMm <= 0)
                return;

            string roomCategory = isEntranceDoor
                ? "Входная"
                : (!string.IsNullOrWhiteSpace(commentValue) ? commentValue.Trim() : "-");

            string key = ApartmentDoorRequirementOption.BuildKey(roomCategory, typeName, widthMm, isEntranceDoor);

            if (result.ContainsKey(key))
                return;

            result.Add(key, new ApartmentDoorRequirementOption
            {
                RoomCategory = roomCategory,
                DoorTypeName2D = typeName,
                WidthMm = widthMm,
                IsEntranceDoor = isEntranceDoor
            });
        }

        private static bool IsEntranceDoorComment(string comment)
        {
            if (string.IsNullOrWhiteSpace(comment))
                return false;

            string normalized = comment
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return string.Equals(normalized, "Входная", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsEntranceProjectDoorSymbol(FamilySymbol symbol)
        {
            if (symbol == null)
                return false;

            string familyName = symbol.Family != null ? symbol.Family.Name ?? "" : "";
            string typeName = symbol.Name ?? "";
            string displayName = BuildDoorTypeDisplayName(symbol);

            return ContainsEntranceDoorName(familyName) ||
                   ContainsEntranceDoorName(typeName) ||
                   ContainsEntranceDoorName(displayName);
        }

        private static bool ContainsEntranceDoorName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .ToLowerInvariant();

            bool hasDoor =
                normalized.IndexOf("двер", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalized.IndexOf("door", StringComparison.OrdinalIgnoreCase) >= 0;

            bool hasEntrance =
                normalized.IndexOf("вход", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalized.IndexOf("entrance", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalized.IndexOf("entry", StringComparison.OrdinalIgnoreCase) >= 0;

            return hasDoor && hasEntrance;
        }

        private static bool Is2DDoorMarker(string familyName, string typeName, string categoryName, bool hasEntranceComment = false)
        {
            bool isGenericModel =
                string.Equals(categoryName, "Обобщенные модели", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(categoryName, "Обобщённые модели", StringComparison.OrdinalIgnoreCase);

            if (!isGenericModel)
                return false;

            return ContainsDoorMarkerText(familyName) ||
                   ContainsDoorMarkerText(typeName) ||
                   hasEntranceComment;
        }

        private static bool ContainsDoorMarkerText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("двер", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("door", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static List<FamilyWindowMarker> CollectWindowMarkersFromApartmentFamily(Document ownerDoc, Family apartmentFamily)
        {
            List<FamilyWindowMarker> windows = new List<FamilyWindowMarker>();
            if (ownerDoc == null || apartmentFamily == null)
                return windows;

            CollectWindowMarkersFromFamilyRecursive(ownerDoc, apartmentFamily, Transform.Identity, windows, 0);

            return windows;
        }

        private static List<FamilyWindowMarker> CollectWindowMarkersFromApartmentInstance(Document ownerDoc, FamilyInstance apartmentInstance)
        {
            List<FamilyWindowMarker> windows = new List<FamilyWindowMarker>();
            if (ownerDoc == null || apartmentInstance == null)
                return windows;

            HashSet<long> visitedInstances = new HashSet<long>();
            CollectPlacedInstanceHelperLinesRecursive(ownerDoc, apartmentInstance, windows, null, null, visitedInstances, 0);

            if (windows.Count == 0 && apartmentInstance.Symbol != null && apartmentInstance.Symbol.Family != null)
            {
                Transform apartmentTransform = apartmentInstance.GetTransform() ?? Transform.Identity;
                foreach (FamilyWindowMarker localMarker in CollectWindowMarkersFromApartmentFamily(ownerDoc, apartmentInstance.Symbol.Family))
                {
                    if (localMarker == null || localMarker.LocalP0 == null || localMarker.LocalP1 == null)
                        continue;

                    windows.Add(new FamilyWindowMarker
                    {
                        LocalP0 = apartmentTransform.OfPoint(localMarker.LocalP0),
                        LocalP1 = apartmentTransform.OfPoint(localMarker.LocalP1)
                    });
                }
            }

            return DeduplicateWindowMarkers(windows);
        }

        private static List<FamilyShaftWallMarker> CollectShaftWallMarkersFromApartmentInstance(Document ownerDoc, FamilyInstance apartmentInstance)
        {
            List<FamilyShaftWallMarker> shafts = new List<FamilyShaftWallMarker>();
            if (ownerDoc == null || apartmentInstance == null)
                return shafts;

            HashSet<long> visitedInstances = new HashSet<long>();
            CollectPlacedInstanceHelperLinesRecursive(ownerDoc, apartmentInstance, null, shafts, null, visitedInstances, 0);

            if (shafts.Count == 0 && apartmentInstance.Symbol != null && apartmentInstance.Symbol.Family != null)
            {
                Transform apartmentTransform = apartmentInstance.GetTransform() ?? Transform.Identity;
                foreach (FamilyShaftWallMarker localMarker in CollectShaftWallMarkersFromApartmentFamily(ownerDoc, apartmentInstance.Symbol.Family))
                {
                    if (localMarker == null || localMarker.ProjectP0 == null || localMarker.ProjectP1 == null)
                        continue;

                    shafts.Add(new FamilyShaftWallMarker
                    {
                        ProjectP0 = apartmentTransform.OfPoint(localMarker.ProjectP0),
                        ProjectP1 = apartmentTransform.OfPoint(localMarker.ProjectP1)
                    });
                }
            }

            return DeduplicateShaftWallMarkers(shafts);
        }

        private static List<FamilyRoomSeparatorMarker> CollectRoomSeparatorMarkersFromApartmentInstance(Document ownerDoc, FamilyInstance apartmentInstance)
        {
            List<FamilyRoomSeparatorMarker> separators = new List<FamilyRoomSeparatorMarker>();
            if (ownerDoc == null || apartmentInstance == null)
                return separators;

            HashSet<long> visitedInstances = new HashSet<long>();
            CollectPlacedInstanceHelperLinesRecursive(ownerDoc, apartmentInstance, null, null, separators, visitedInstances, 0);

            if (separators.Count == 0 && apartmentInstance.Symbol != null && apartmentInstance.Symbol.Family != null)
            {
                Transform apartmentTransform = apartmentInstance.GetTransform() ?? Transform.Identity;
                foreach (FamilyRoomSeparatorMarker localMarker in CollectRoomSeparatorMarkersFromApartmentFamily(ownerDoc, apartmentInstance.Symbol.Family))
                {
                    if (localMarker == null || localMarker.ProjectP0 == null || localMarker.ProjectP1 == null)
                        continue;

                    separators.Add(new FamilyRoomSeparatorMarker
                    {
                        ProjectP0 = apartmentTransform.OfPoint(localMarker.ProjectP0),
                        ProjectP1 = apartmentTransform.OfPoint(localMarker.ProjectP1)
                    });
                }
            }

            return DeduplicateRoomSeparatorMarkers(separators);
        }

        private static List<FamilyLoggiaWallMarker> CollectLoggiaWallMarkersFromApartmentInstance(Document ownerDoc, FamilyInstance apartmentInstance)
        {
            List<FamilyLoggiaWallMarker> loggias = new List<FamilyLoggiaWallMarker>();
            if (ownerDoc == null || apartmentInstance == null)
                return loggias;

            if (!IsLoggiaEnabled(apartmentInstance))
                return loggias;

            HashSet<long> visitedInstances = new HashSet<long>();
            CollectPlacedLoggiaWallMarkersRecursive(ownerDoc, apartmentInstance, loggias, visitedInstances, 0);

            return DeduplicateLoggiaWallMarkers(loggias);
        }

        private static void CollectPlacedLoggiaWallMarkersRecursive(Document ownerDoc, FamilyInstance instance,
            List<FamilyLoggiaWallMarker> loggias, HashSet<long> visitedInstanceIds, int depth)
        {
            if (ownerDoc == null || instance == null || loggias == null || visitedInstanceIds == null)
                return;

            if (depth > 12)
                return;

            long id = IDHelper.ElIdValue(instance.Id);
            if (visitedInstanceIds.Contains(id))
                return;

            visitedInstanceIds.Add(id);

            AddLoggiaHelperLinesFromPlacedInstance(ownerDoc, instance, loggias);

            ICollection<ElementId> subIds = null;
            try
            {
                subIds = instance.GetSubComponentIds();
            }
            catch
            {
                subIds = null;
            }

            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = ownerDoc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                CollectPlacedLoggiaWallMarkersRecursive(ownerDoc, subFi, loggias, visitedInstanceIds, depth + 1);
            }
        }

        private static void AddLoggiaHelperLinesFromPlacedInstance(Document doc, FamilyInstance instance, List<FamilyLoggiaWallMarker> loggias)
        {
            if (doc == null || instance == null || loggias == null)
                return;

            if (IsLoggiaWallHelperInstance(instance))
                AddLongestLoggiaWallMarkerFromInstanceGeometry(doc, instance, loggias);
            else
                AddStyledLoggiaWallMarkersFromInstanceGeometry(doc, instance, loggias);
        }

        private static void CollectPlacedInstanceHelperLinesRecursive(Document ownerDoc, FamilyInstance instance,
            List<FamilyWindowMarker> windows, List<FamilyShaftWallMarker> shafts, List<FamilyRoomSeparatorMarker> separators,
            HashSet<long> visitedInstanceIds, int depth)
        {
            if (ownerDoc == null || instance == null || visitedInstanceIds == null)
                return;

            if (depth > 12)
                return;

            long id = IDHelper.ElIdValue(instance.Id);
            if (visitedInstanceIds.Contains(id))
                return;

            visitedInstanceIds.Add(id);

            AddHelperLinesFromPlacedInstance(ownerDoc, instance, windows, shafts, separators);

            ICollection<ElementId> subIds = null;
            try
            {
                subIds = instance.GetSubComponentIds();
            }
            catch
            {
                subIds = null;
            }

            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = ownerDoc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                CollectPlacedInstanceHelperLinesRecursive(ownerDoc, subFi, windows, shafts, separators, visitedInstanceIds, depth + 1);
            }
        }

        private static void AddHelperLinesFromPlacedInstance(Document doc, FamilyInstance instance,
            List<FamilyWindowMarker> windows, List<FamilyShaftWallMarker> shafts, List<FamilyRoomSeparatorMarker> separators)
        {
            if (doc == null || instance == null)
                return;

            bool explicitWindow = IsWindowHelperInstance(instance);
            bool explicitShaft = IsShaftWallHelperInstance(instance);
            bool explicitSeparator = IsRoomSeparatorHelperInstance(instance);

            if (windows != null)
            {
                if (explicitWindow)
                    AddLongestWindowMarkerFromInstanceGeometry(doc, instance, windows);
                else
                    AddStyledWindowMarkersFromInstanceGeometry(doc, instance, windows);
            }

            if (shafts != null)
            {
                if (explicitShaft)
                    AddLongestShaftWallMarkerFromInstanceGeometry(doc, instance, shafts);
                else
                    AddStyledShaftWallMarkersFromInstanceGeometry(doc, instance, shafts);
            }

            if (separators != null)
            {
                if (explicitSeparator)
                    AddLongestRoomSeparatorMarkerFromInstanceGeometry(doc, instance, separators);
                else
                    AddStyledRoomSeparatorMarkersFromInstanceGeometry(doc, instance, separators);
            }
        }

        private static void AddLongestWindowMarkerFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyWindowMarker> windows)
        {
            HelperLineCandidate best = GetBestHelperLineCandidate(doc, instance, IsWindowHelperName, true);
            if (best == null)
                return;

            windows.Add(new FamilyWindowMarker
            {
                LocalP0 = best.P0,
                LocalP1 = best.P1
            });
        }

        private static void AddStyledWindowMarkersFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyWindowMarker> windows)
        {
            foreach (HelperLineCandidate line in CollectHelperLineCandidates(doc, instance, IsWindowHelperName, false))
            {
                windows.Add(new FamilyWindowMarker
                {
                    LocalP0 = line.P0,
                    LocalP1 = line.P1
                });
            }
        }

        private static void AddLongestShaftWallMarkerFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyShaftWallMarker> shafts)
        {
            HelperLineCandidate best = GetBestHelperLineCandidate(doc, instance, IsShaftWallHelperName, true);
            if (best == null)
                return;

            shafts.Add(new FamilyShaftWallMarker
            {
                ProjectP0 = best.P0,
                ProjectP1 = best.P1
            });
        }

        private static void AddStyledShaftWallMarkersFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyShaftWallMarker> shafts)
        {
            foreach (HelperLineCandidate line in CollectHelperLineCandidates(doc, instance, IsShaftWallHelperName, false))
            {
                shafts.Add(new FamilyShaftWallMarker
                {
                    ProjectP0 = line.P0,
                    ProjectP1 = line.P1
                });
            }
        }

        private static void AddLongestRoomSeparatorMarkerFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyRoomSeparatorMarker> separators)
        {
            HelperLineCandidate best = GetBestHelperLineCandidate(doc, instance, IsRoomSeparatorHelperName, true);
            if (best == null)
                return;

            separators.Add(new FamilyRoomSeparatorMarker
            {
                ProjectP0 = best.P0,
                ProjectP1 = best.P1
            });
        }

        private static void AddStyledRoomSeparatorMarkersFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyRoomSeparatorMarker> separators)
        {
            foreach (HelperLineCandidate line in CollectHelperLineCandidates(doc, instance, IsRoomSeparatorHelperName, false))
            {
                separators.Add(new FamilyRoomSeparatorMarker
                {
                    ProjectP0 = line.P0,
                    ProjectP1 = line.P1
                });
            }
        }

        private static void AddLongestLoggiaWallMarkerFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyLoggiaWallMarker> loggias)
        {
            HelperLineCandidate best = GetBestHelperLineCandidate(doc, instance, IsLoggiaWallHelperName, true);
            if (best == null)
                return;

            loggias.Add(new FamilyLoggiaWallMarker
            {
                ProjectP0 = best.P0,
                ProjectP1 = best.P1
            });
        }

        private static void AddStyledLoggiaWallMarkersFromInstanceGeometry(Document doc, FamilyInstance instance, List<FamilyLoggiaWallMarker> loggias)
        {
            foreach (HelperLineCandidate line in CollectHelperLineCandidates(doc, instance, IsLoggiaWallHelperName, false))
            {
                loggias.Add(new FamilyLoggiaWallMarker
                {
                    ProjectP0 = line.P0,
                    ProjectP1 = line.P1
                });
            }
        }

        private static HelperLineCandidate GetBestHelperLineCandidate(Document doc, FamilyInstance instance, Func<string, bool> styleNamePredicate,
            bool allowAnyStyle)
        {
            List<HelperLineCandidate> candidates = CollectHelperLineCandidates(doc, instance, styleNamePredicate, allowAnyStyle);
            if (candidates.Count == 0)
                return null;

            List<HelperLineCandidate> styled = candidates.Where(x => x != null && x.StyleMatched).ToList();
            if (styled.Count > 0)
                candidates = styled;

            return candidates
                .Where(x => x != null && x.P0 != null && x.P1 != null)
                .OrderByDescending(x => Distance2D(x.P0, x.P1))
                .FirstOrDefault();
        }

        private static List<HelperLineCandidate> CollectHelperLineCandidates(Document doc, FamilyInstance instance, Func<string, bool> styleNamePredicate,
            bool allowAnyStyle)
        {
            List<HelperLineCandidate> result = new List<HelperLineCandidate>();
            if (doc == null || instance == null)
                return result;

            Options options = new Options
            {
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geometry = null;
            try
            {
                geometry = instance.get_Geometry(options);
            }
            catch
            {
                geometry = null;
            }

            CollectHelperLineCandidatesFromGeometry(doc, geometry, Transform.Identity, styleNamePredicate, allowAnyStyle, result);

            return result;
        }

        private static void CollectHelperLineCandidatesFromGeometry(Document doc, GeometryElement geometry, Transform transform,
            Func<string, bool> styleNamePredicate, bool allowAnyStyle, List<HelperLineCandidate> result)
        {
            if (doc == null || geometry == null || transform == null || result == null)
                return;

            foreach (GeometryObject obj in geometry)
            {
                if (obj == null)
                    continue;

                GeometryInstance geometryInstance = obj as GeometryInstance;
                if (geometryInstance != null)
                {
                    Transform nestedTransform = transform;
                    try
                    {
                        if (geometryInstance.Transform != null)
                            nestedTransform = transform.Multiply(geometryInstance.Transform);
                    }
                    catch
                    {
                        nestedTransform = transform;
                    }

                    GeometryElement symbolGeometry = null;
                    try
                    {
                        symbolGeometry = geometryInstance.GetSymbolGeometry();
                    }
                    catch
                    {
                        symbolGeometry = null;
                    }

                    CollectHelperLineCandidatesFromGeometry(doc, symbolGeometry, nestedTransform, styleNamePredicate, allowAnyStyle, result);
                    continue;
                }

                Curve curve = obj as Curve;
                if (curve == null || !curve.IsBound)
                    continue;

                bool styleMatched = GeometryObjectMatchesName(doc, obj, styleNamePredicate);
                if (!allowAnyStyle && !styleMatched)
                    continue;

                XYZ p0;
                XYZ p1;

                try
                {
                    p0 = transform.OfPoint(curve.GetEndPoint(0));
                    p1 = transform.OfPoint(curve.GetEndPoint(1));
                }
                catch
                {
                    continue;
                }

                if (p0 == null || p1 == null || Distance2D(p0, p1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                result.Add(new HelperLineCandidate
                {
                    P0 = p0,
                    P1 = p1,
                    StyleMatched = styleMatched
                });
            }
        }

        private static bool GeometryObjectMatchesName(Document doc, GeometryObject obj, Func<string, bool> namePredicate)
        {
            if (doc == null || obj == null || namePredicate == null)
                return false;

            ElementId graphicsStyleId = ElementId.InvalidElementId;
            try
            {
                graphicsStyleId = obj.GraphicsStyleId;
            }
            catch
            {
                graphicsStyleId = ElementId.InvalidElementId;
            }

            if (graphicsStyleId == ElementId.InvalidElementId)
                return false;

            GraphicsStyle graphicsStyle = doc.GetElement(graphicsStyleId) as GraphicsStyle;
            if (graphicsStyle == null)
                return false;

            if (namePredicate(graphicsStyle.Name))
                return true;

            Category category = graphicsStyle.GraphicsStyleCategory;
            return category != null && namePredicate(category.Name);
        }

        private static bool IsWindowHelperInstance(FamilyInstance instance)
        {
            return HelperInstanceMatchesName(instance, IsWindowHelperName);
        }

        private static bool IsShaftWallHelperInstance(FamilyInstance instance)
        {
            return HelperInstanceMatchesName(instance, IsShaftWallHelperName);
        }

        private static bool IsRoomSeparatorHelperInstance(FamilyInstance instance)
        {
            return HelperInstanceMatchesName(instance, IsRoomSeparatorHelperName);
        }

        private static bool IsLoggiaWallHelperInstance(FamilyInstance instance)
        {
            return HelperInstanceMatchesName(instance, IsLoggiaWallHelperName);
        }

        private static bool IsLoggiaEnabled(FamilyInstance apartmentInstance)
        {
            bool value;
            return TryGetYesNoParamFromElement(apartmentInstance, out value, "Лоджия", "Loggia") && value;
        }

        private static bool HelperInstanceMatchesName(FamilyInstance instance, Func<string, bool> namePredicate)
        {
            if (instance == null || namePredicate == null)
                return false;

            List<string> values = new List<string>();

            values.Add(GetStringParameterValue(instance, "Подкатегория", "Subcategory", "Субкатегория"));

            if (instance.Symbol != null)
            {
                values.Add(instance.Symbol.Name);
                values.Add(GetStringParameterValue(instance.Symbol, "Подкатегория", "Subcategory", "Субкатегория"));

                if (instance.Symbol.Family != null)
                    values.Add(instance.Symbol.Family.Name);
            }

            if (instance.Category != null)
                values.Add(instance.Category.Name);

            return values.Any(x => namePredicate(x));
        }

        private static bool TryGetYesNoParamFromElement(Element element, out bool value, params string[] paramNames)
        {
            value = false;

            if (element == null || paramNames == null || paramNames.Length == 0)
                return false;

            foreach (string paramName in paramNames)
            {
                if (string.IsNullOrWhiteSpace(paramName))
                    continue;

                Parameter p = element.LookupParameter(paramName);
                if (TryGetYesNoParameterValue(p, out value))
                    return true;
            }

            return false;
        }

        private static bool TryGetYesNoParameterValue(Parameter parameter, out bool value)
        {
            value = false;

            if (parameter == null)
                return false;

            try
            {
                if (parameter.StorageType == StorageType.Integer)
                {
                    value = parameter.AsInteger() != 0;
                    return true;
                }

                if (parameter.StorageType == StorageType.String)
                {
                    string s = parameter.AsString();
                    if (string.IsNullOrWhiteSpace(s))
                        return false;

                    s = s.Trim();
                    if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "да", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                    {
                        value = true;
                        return true;
                    }

                    if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "нет", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                    {
                        value = false;
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static string GetStringParameterValue(Element element, params string[] paramNames)
        {
            if (element == null || paramNames == null || paramNames.Length == 0)
                return null;

            foreach (string paramName in paramNames)
            {
                if (string.IsNullOrWhiteSpace(paramName))
                    continue;

                Parameter p = element.LookupParameter(paramName);
                string value = GetStringParameterValue(p);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            Document doc = element.Document;
            Element typeElem = null;
            try
            {
                ElementId typeId = element.GetTypeId();
                if (doc != null && typeId != ElementId.InvalidElementId)
                    typeElem = doc.GetElement(typeId);
            }
            catch
            {
                typeElem = null;
            }

            if (typeElem == null)
                return null;

            foreach (string paramName in paramNames)
            {
                if (string.IsNullOrWhiteSpace(paramName))
                    continue;

                Parameter p = typeElem.LookupParameter(paramName);
                string value = GetStringParameterValue(p);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return null;
        }

        private static string GetStringParameterValue(Parameter p)
        {
            if (p == null)
                return null;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                string valueString = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                    return valueString;
            }
            catch
            {
            }

            return null;
        }

        private static List<FamilyWindowMarker> DeduplicateWindowMarkers(List<FamilyWindowMarker> source)
        {
            List<FamilyWindowMarker> result = new List<FamilyWindowMarker>();
            if (source == null)
                return result;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyWindowMarker marker in source)
            {
                if (marker == null || marker.LocalP0 == null || marker.LocalP1 == null)
                    continue;

                string key = BuildLineMarkerKey(marker.LocalP0, marker.LocalP1);
                if (keys.Contains(key))
                    continue;

                keys.Add(key);
                result.Add(marker);
            }

            return result;
        }

        private static List<FamilyShaftWallMarker> DeduplicateShaftWallMarkers(List<FamilyShaftWallMarker> source)
        {
            List<FamilyShaftWallMarker> result = new List<FamilyShaftWallMarker>();
            if (source == null)
                return result;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyShaftWallMarker marker in source)
            {
                if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                    continue;

                string key = BuildLineMarkerKey(marker.ProjectP0, marker.ProjectP1);
                if (keys.Contains(key))
                    continue;

                keys.Add(key);
                result.Add(marker);
            }

            return result;
        }

        private static List<FamilyLoggiaWallMarker> DeduplicateLoggiaWallMarkers(List<FamilyLoggiaWallMarker> source)
        {
            List<FamilyLoggiaWallMarker> result = new List<FamilyLoggiaWallMarker>();
            if (source == null)
                return result;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyLoggiaWallMarker marker in source)
            {
                if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                    continue;

                string key = BuildLineMarkerKey(marker.ProjectP0, marker.ProjectP1);
                if (keys.Contains(key))
                    continue;

                keys.Add(key);
                result.Add(marker);
            }

            return result;
        }

        private static List<FamilyRoomSeparatorMarker> DeduplicateRoomSeparatorMarkers(List<FamilyRoomSeparatorMarker> source)
        {
            List<FamilyRoomSeparatorMarker> result = new List<FamilyRoomSeparatorMarker>();
            if (source == null)
                return result;

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyRoomSeparatorMarker marker in source)
            {
                if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                    continue;

                string key = BuildLineMarkerKey(marker.ProjectP0, marker.ProjectP1);
                if (keys.Contains(key))
                    continue;

                keys.Add(key);
                result.Add(marker);
            }

            return result;
        }

        private static string BuildLineMarkerKey(XYZ p0, XYZ p1)
        {
            string a = BuildPointMarkerKey(p0);
            string b = BuildPointMarkerKey(p1);

            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase) <= 0
                ? a + "|" + b
                : b + "|" + a;
        }

        private static string BuildPointMarkerKey(XYZ p)
        {
            if (p == null)
                return "";

            return Math.Round(p.X, 6) + ";" + Math.Round(p.Y, 6) + ";" + Math.Round(p.Z, 6);
        }

        private static void CollectWindowMarkersFromFamilyRecursive(Document ownerDoc, Family family, Transform accumulatedLocalTransform,
            List<FamilyWindowMarker> windows, int depth)
        {
            if (ownerDoc == null || family == null || accumulatedLocalTransform == null || windows == null)
                return;

            if (depth > 12)
                return;

            Document familyDoc = null;

            try
            {
                familyDoc = ownerDoc.EditFamily(family);

                CollectWindowHelperLines(familyDoc, accumulatedLocalTransform, windows);

                List<FamilyInstance> nestedInstances = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (FamilyInstance fi in nestedInstances)
                {
                    if (fi == null || fi.Symbol == null || fi.Symbol.Family == null)
                        continue;

                    Transform localTransform = fi.GetTransform();
                    if (localTransform == null)
                        localTransform = Transform.Identity;

                    Transform currentLocalTransform = accumulatedLocalTransform.Multiply(localTransform);
                    CollectWindowMarkersFromFamilyRecursive(familyDoc, fi.Symbol.Family, currentLocalTransform, windows, depth + 1);
                }
            }
            catch
            {
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static List<FamilyShaftWallMarker> CollectShaftWallMarkersFromApartmentFamily(Document ownerDoc, Family apartmentFamily)
        {
            List<FamilyShaftWallMarker> shafts = new List<FamilyShaftWallMarker>();
            if (ownerDoc == null || apartmentFamily == null)
                return shafts;

            CollectShaftWallMarkersFromFamilyRecursive(ownerDoc, apartmentFamily, Transform.Identity, shafts, 0);

            return shafts;
        }

        private static List<FamilyRoomSeparatorMarker> CollectRoomSeparatorMarkersFromApartmentFamily(Document ownerDoc, Family apartmentFamily)
        {
            List<FamilyRoomSeparatorMarker> separators = new List<FamilyRoomSeparatorMarker>();
            if (ownerDoc == null || apartmentFamily == null)
                return separators;

            CollectRoomSeparatorMarkersFromFamilyRecursive(ownerDoc, apartmentFamily, Transform.Identity, separators, 0);

            return separators;
        }

        private static void CollectShaftWallMarkersFromFamilyRecursive(Document ownerDoc, Family family, Transform accumulatedLocalTransform,
            List<FamilyShaftWallMarker> shafts, int depth)
        {
            if (ownerDoc == null || family == null || accumulatedLocalTransform == null || shafts == null)
                return;

            if (depth > 12)
                return;

            Document familyDoc = null;

            try
            {
                familyDoc = ownerDoc.EditFamily(family);

                CollectShaftWallHelperLines(familyDoc, accumulatedLocalTransform, shafts);

                List<FamilyInstance> nestedInstances = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (FamilyInstance fi in nestedInstances)
                {
                    if (fi == null || fi.Symbol == null || fi.Symbol.Family == null)
                        continue;

                    Transform localTransform = fi.GetTransform();
                    if (localTransform == null)
                        localTransform = Transform.Identity;

                    Transform currentLocalTransform = accumulatedLocalTransform.Multiply(localTransform);
                    CollectShaftWallMarkersFromFamilyRecursive(familyDoc, fi.Symbol.Family, currentLocalTransform, shafts, depth + 1);
                }
            }
            catch
            {
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void CollectRoomSeparatorMarkersFromFamilyRecursive(Document ownerDoc, Family family, Transform accumulatedLocalTransform,
            List<FamilyRoomSeparatorMarker> separators, int depth)
        {
            if (ownerDoc == null || family == null || accumulatedLocalTransform == null || separators == null)
                return;

            if (depth > 12)
                return;

            Document familyDoc = null;

            try
            {
                familyDoc = ownerDoc.EditFamily(family);

                CollectRoomSeparatorHelperLines(familyDoc, accumulatedLocalTransform, separators);

                List<FamilyInstance> nestedInstances = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (FamilyInstance fi in nestedInstances)
                {
                    if (fi == null || fi.Symbol == null || fi.Symbol.Family == null)
                        continue;

                    Transform localTransform = fi.GetTransform();
                    if (localTransform == null)
                        localTransform = Transform.Identity;

                    Transform currentLocalTransform = accumulatedLocalTransform.Multiply(localTransform);
                    CollectRoomSeparatorMarkersFromFamilyRecursive(familyDoc, fi.Symbol.Family, currentLocalTransform, separators, depth + 1);
                }
            }
            catch
            {
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void CollectWindowHelperLines(Document familyDoc, Transform accumulatedLocalTransform, List<FamilyWindowMarker> windows)
        {
            if (familyDoc == null || accumulatedLocalTransform == null || windows == null)
                return;

            List<CurveElement> curves = new FilteredElementCollector(familyDoc)
                .OfClass(typeof(CurveElement))
                .Cast<CurveElement>()
                .ToList();

            foreach (CurveElement curveElement in curves)
            {
                if (curveElement == null || !IsWindowHelperCurve(curveElement))
                    continue;

                Curve curve = null;
                try
                {
                    curve = curveElement.GeometryCurve;
                }
                catch
                {
                    curve = null;
                }

                if (curve == null || !curve.IsBound)
                    continue;

                XYZ p0;
                XYZ p1;

                try
                {
                    p0 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(0));
                    p1 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(1));
                }
                catch
                {
                    continue;
                }

                if (p0 == null || p1 == null || Distance2D(p0, p1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                windows.Add(new FamilyWindowMarker
                {
                    LocalP0 = p0,
                    LocalP1 = p1
                });
            }
        }

        private static bool IsWindowHelperCurve(CurveElement curveElement)
        {
            if (curveElement == null)
                return false;

            Element lineStyle = null;
            try
            {
                lineStyle = curveElement.LineStyle;
            }
            catch
            {
                lineStyle = null;
            }

            GraphicsStyle graphicsStyle = lineStyle as GraphicsStyle;
            Category styleCategory = graphicsStyle != null ? graphicsStyle.GraphicsStyleCategory : null;

            if (IsWindowHelperName(styleCategory != null ? styleCategory.Name : null))
                return true;

            if (IsWindowHelperName(curveElement.Category != null ? curveElement.Category.Name : null))
                return true;

            return false;
        }

        private static bool IsWindowHelperName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("окно", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void CollectShaftWallHelperLines(Document familyDoc, Transform accumulatedLocalTransform, List<FamilyShaftWallMarker> shafts)
        {
            if (familyDoc == null || accumulatedLocalTransform == null || shafts == null)
                return;

            List<CurveElement> curves = new FilteredElementCollector(familyDoc)
                .OfClass(typeof(CurveElement))
                .Cast<CurveElement>()
                .ToList();

            foreach (CurveElement curveElement in curves)
            {
                if (curveElement == null || !IsShaftWallHelperCurve(curveElement))
                    continue;

                Curve curve = null;
                try
                {
                    curve = curveElement.GeometryCurve;
                }
                catch
                {
                    curve = null;
                }

                if (curve == null || !curve.IsBound)
                    continue;

                XYZ p0;
                XYZ p1;

                try
                {
                    p0 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(0));
                    p1 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(1));
                }
                catch
                {
                    continue;
                }

                if (p0 == null || p1 == null || Distance2D(p0, p1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                shafts.Add(new FamilyShaftWallMarker
                {
                    ProjectP0 = p0,
                    ProjectP1 = p1
                });
            }
        }

        private static void CollectRoomSeparatorHelperLines(Document familyDoc, Transform accumulatedLocalTransform, List<FamilyRoomSeparatorMarker> separators)
        {
            if (familyDoc == null || accumulatedLocalTransform == null || separators == null)
                return;

            List<CurveElement> curves = new FilteredElementCollector(familyDoc)
                .OfClass(typeof(CurveElement))
                .Cast<CurveElement>()
                .ToList();

            foreach (CurveElement curveElement in curves)
            {
                if (curveElement == null || !IsRoomSeparatorHelperCurve(curveElement))
                    continue;

                Curve curve = null;
                try
                {
                    curve = curveElement.GeometryCurve;
                }
                catch
                {
                    curve = null;
                }

                if (curve == null || !curve.IsBound)
                    continue;

                XYZ p0;
                XYZ p1;

                try
                {
                    p0 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(0));
                    p1 = accumulatedLocalTransform.OfPoint(curve.GetEndPoint(1));
                }
                catch
                {
                    continue;
                }

                if (p0 == null || p1 == null || Distance2D(p0, p1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                separators.Add(new FamilyRoomSeparatorMarker
                {
                    ProjectP0 = p0,
                    ProjectP1 = p1
                });
            }
        }

        private static bool IsShaftWallHelperCurve(CurveElement curveElement)
        {
            if (curveElement == null)
                return false;

            Element lineStyle = null;
            try
            {
                lineStyle = curveElement.LineStyle;
            }
            catch
            {
                lineStyle = null;
            }

            GraphicsStyle graphicsStyle = lineStyle as GraphicsStyle;
            Category styleCategory = graphicsStyle != null ? graphicsStyle.GraphicsStyleCategory : null;

            if (IsShaftWallHelperName(styleCategory != null ? styleCategory.Name : null))
                return true;

            if (IsShaftWallHelperName(curveElement.Category != null ? curveElement.Category.Name : null))
                return true;

            return false;
        }

        private static bool IsRoomSeparatorHelperCurve(CurveElement curveElement)
        {
            if (curveElement == null)
                return false;

            Element lineStyle = null;
            try
            {
                lineStyle = curveElement.LineStyle;
            }
            catch
            {
                lineStyle = null;
            }

            GraphicsStyle graphicsStyle = lineStyle as GraphicsStyle;
            Category styleCategory = graphicsStyle != null ? graphicsStyle.GraphicsStyleCategory : null;

            if (IsRoomSeparatorHelperName(styleCategory != null ? styleCategory.Name : null))
                return true;

            if (IsRoomSeparatorHelperName(curveElement.Category != null ? curveElement.Category.Name : null))
                return true;

            return false;
        }

        private static bool IsShaftWallHelperName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("шахта", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsRoomSeparatorHelperName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("разделитель", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsLoggiaWallHelperName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = value
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("лоджия", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("loggia", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryGetDoorWidthMmFrom2DTypeName(string typeName, out int widthMm)
        {
            widthMm = 0;

            if (string.IsNullOrWhiteSpace(typeName))
                return false;

            string normalized = typeName.Trim();

            int parsed;
            if (int.TryParse(normalized, out parsed))
            {
                widthMm = parsed;
                return widthMm > 0;
            }

            List<int> numbers = Regex.Matches(normalized, @"\d+")
                .Cast<Match>()
                .Select(x =>
                {
                    int value;
                    return int.TryParse(x.Value, out value) ? value : 0;
                })
                .Where(x => x > 0)
                .ToList();

            if (numbers.Count == 0)
                return false;

            if (numbers.Count == 1)
            {
                widthMm = numbers[0];
                return widthMm > 0;
            }

            int preferredWidth = numbers
                .Where(x => x >= 300 && x <= 2000)
                .OrderBy(x => x)
                .FirstOrDefault();

            widthMm = preferredWidth > 0
                ? preferredWidth
                : numbers[0];

            return widthMm > 0;
        }

        private static bool TryGetDoorWidthMmFrom2DMarker(FamilyInstance doorFi, string typeName, out int widthMm)
        {
            int widthFromTypeName;
            bool hasWidthFromTypeName = TryGetDoorWidthMmFrom2DTypeName(typeName, out widthFromTypeName) && widthFromTypeName > 0;
            if (hasWidthFromTypeName && widthFromTypeName <= 2000)
            {
                widthMm = widthFromTypeName;
                return true;
            }

            widthMm = 0;

            double widthInternal;
            if (TryGetLengthParamFromElementOrType(
                doorFi,
                out widthInternal,
                "Ширина",
                "Width",
                "КП_Ширина",
                "ADSK_Размер_Ширина"))
            {
                widthMm = (int)Math.Round(IDHelper.ConvertInternalToMm(widthInternal));
                return widthMm > 0;
            }

            if (hasWidthFromTypeName)
            {
                widthMm = widthFromTypeName;
                return true;
            }

            return false;
        }

        private Dictionary<string, List<string>> BuildDoorTypeOptionsByRequirementForPlan(Document doc, List<ApartmentDoorRequirementOption> requirements)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            if (requirements == null || requirements.Count == 0)
                return result;

            List<FamilySymbol> doorTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            foreach (ApartmentDoorRequirementOption requirement in requirements)
            {
                if (requirement == null || string.IsNullOrWhiteSpace(requirement.DoorTypeName2D) || requirement.WidthMm <= 0)
                    continue;

                List<string> matched = new List<string>();
                List<string> matchedByWidth = new List<string>();

                foreach (FamilySymbol symbol in doorTypes)
                {
                    int widthMm;
                    if (!TryGetProjectDoorTypeWidthMm(symbol, out widthMm))
                        continue;

                    if (widthMm != requirement.WidthMm)
                        continue;

                    string displayName = BuildDoorTypeDisplayName(symbol);
                    if (!string.IsNullOrWhiteSpace(displayName))
                    {
                        matchedByWidth.Add(displayName);

                        bool isEntranceSymbol = IsEntranceProjectDoorSymbol(symbol);
                        if (requirement.IsEntranceDoor)
                        {
                            if (isEntranceSymbol)
                                matched.Add(displayName);
                        }
                        else if (!isEntranceSymbol)
                        {
                            matched.Add(displayName);
                        }
                    }
                }

                if (requirement.IsEntranceDoor && matched.Count == 0)
                    matched = matchedByWidth;

                result[requirement.Key] = matched
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();
            }

            return result;
        }

        private static bool TryGetProjectDoorTypeWidthMm(FamilySymbol symbol, out int widthMm)
        {
            widthMm = 0;
            if (symbol == null)
                return false;

            double widthInternal;
            if (TryGetLengthParamFromElementOrType(symbol, out widthInternal, "Ширина", "Width"))
            {
                widthMm = (int)Math.Round(IDHelper.ConvertInternalToMm(widthInternal));
                return widthMm > 0;
            }

            Parameter builtInParam = symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            if (builtInParam != null && builtInParam.StorageType == StorageType.Double)
            {
                widthMm = (int)Math.Round(IDHelper.ConvertInternalToMm(builtInParam.AsDouble()));
                return widthMm > 0;
            }

            return false;
        }

        private static bool TryGetLengthParamFromElementOrType(Element e, out double valueInternal, params string[] paramNames)
        {
            valueInternal = 0;
            if (e == null)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    valueInternal = p.AsDouble();
                    return true;
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        valueInternal = p.AsDouble();
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetAreaParamFromElementOrType(Element e, out double valueInternal, params string[] paramNames)
        {
            valueInternal = 0;
            if (e == null)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    valueInternal = p.AsDouble();
                    return true;
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        valueInternal = p.AsDouble();
                        return true;
                    }
                }
            }

            return false;
        }

        private static string BuildDoorTypeDisplayName(FamilySymbol symbol)
        {
            return BuildFamilySymbolDisplayName(symbol);
        }

        private static string BuildWindowTypeDisplayName(FamilySymbol symbol)
        {
            return BuildFamilySymbolDisplayName(symbol);
        }

        private static string BuildFamilySymbolDisplayName(FamilySymbol symbol)
        {
            if (symbol == null)
                return null;

            string familyName = symbol.Family != null ? symbol.Family.Name ?? "" : "";
            string typeName = symbol.Name ?? "";

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
                return familyName + " - " + typeName;

            if (!string.IsNullOrWhiteSpace(typeName))
                return typeName;

            return familyName;
        }

        private static FamilySymbol FindWindowSymbolByDisplayName(Document doc, string displayName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(displayName))
                return null;

            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(x => string.Equals(BuildWindowTypeDisplayName(x), displayName, StringComparison.OrdinalIgnoreCase));
        }

        private static FamilySymbol FindDoorSymbolByDisplayNameAndWidth(Document doc, string displayName, int widthMm, bool? isEntranceDoor = null)
        {
            if (doc == null || string.IsNullOrWhiteSpace(displayName))
                return null;

            List<FamilySymbol> doorTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            foreach (FamilySymbol symbol in doorTypes)
            {
                string currentName = BuildDoorTypeDisplayName(symbol);
                if (!string.Equals(currentName, displayName, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (isEntranceDoor.HasValue && !isEntranceDoor.Value && IsEntranceProjectDoorSymbol(symbol))
                    continue;

                int symbolWidthMm;
                if (!TryGetProjectDoorTypeWidthMm(symbol, out symbolWidthMm))
                    continue;

                if (symbolWidthMm == widthMm)
                    return symbol;
            }

            return null;
        }

        private bool ValidatePresetBeforeConvertTo3D(Document doc, ApartmentPresetData preset, out string validationMessage, out bool isPresetDataStale)
        {
            validationMessage = "";
            isPresetDataStale = false;

            ApartmentPresetData effectivePreset = preset ?? new ApartmentPresetData
            {
                SelectedPlanName = "",
                SelectedPlanModelSignature = "",
                LowerConstraint = "",
                UpperConstraint = "Неприсоединённая",
                BaseOffset = 0,
                WallHeight = 3000,
                WallTypeByThickness = new Dictionary<int, string>(),
                WindowType = "Не выбрано",
                WindowSillHeight = 900,
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = new Dictionary<string, string>(),
                FamilyPostProcessAction = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            List<string> notFilled = new List<string>();
            List<string> otherProblems = new List<string>();

            if (effectivePreset.WallHeight <= 0)
                notFilled.Add("Неприсоединённая высота стены");

            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
            if (targetPlan == null)
            {
                otherProblems.Add("Не удалось определить план для построения стен.");
            }
            else
            {
                string savedSignature = effectivePreset.SelectedPlanModelSignature;
                string currentSignature = BuildApartmentPlanModelSignature(doc, targetPlan);
                if (!string.IsNullOrWhiteSpace(savedSignature) &&
                    !string.Equals(savedSignature, currentSignature, StringComparison.Ordinal))
                {
                    isPresetDataStale = true;
                    otherProblems.Add("Данные плана устарели: изменился состав 2D-семейств квартир или параметр 'Лоджия'. Нажмите 'Обновить данные'.");
                }

                List<int> requiredThicknesses = BuildWallThicknessesForPlan(doc, targetPlan);

                if (requiredThicknesses.Count == 0)
                {
                    otherProblems.Add("На выбранном плане не найдено ни одного размещённого 2D-семейства квартиры.");
                }
                else
                {
                    foreach (int thickness in requiredThicknesses.OrderBy(x => x))
                    {
                        string selectedWallType = GetSelectedWallTypeNameForThickness(effectivePreset, thickness);

                        if (string.IsNullOrWhiteSpace(selectedWallType) ||
                            string.Equals(selectedWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        {
                            notFilled.Add("Стена (" + thickness + ")");
                            continue;
                        }

                        WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallType, thickness);
                        if (matchedWallType == null)
                        {
                            otherProblems.Add(
                                "Для толщины " + thickness + " мм выбран тип стены '" + selectedWallType + "', но такой тип не найден в проекте.");
                        }
                    }
                }

                List<ApartmentDoorRequirementOption> doorRequirements = BuildDoorRequirementsForPlan(doc, targetPlan);

                bool hasWindowMarkers = HasWindowMarkersForPlan(doc, targetPlan);
                if (hasWindowMarkers)
                {
                    if (effectivePreset.WindowSillHeight <= 0)
                        notFilled.Add("Высота нижнего бруса окна");

                    if (string.IsNullOrWhiteSpace(effectivePreset.WindowType) ||
                        string.Equals(effectivePreset.WindowType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                    {
                        notFilled.Add("Тип окна");
                    }
                    else
                    {
                        FamilySymbol matchedWindowSymbol = FindWindowSymbolByDisplayName(doc, effectivePreset.WindowType);
                        if (matchedWindowSymbol == null)
                        {
                            otherProblems.Add("Выбран тип окна '" + effectivePreset.WindowType + "', но такой тип не найден в проекте.");
                        }
                    }
                }

                bool hasShaftWallMarkers = HasShaftWallMarkersForPlan(doc, targetPlan);
                if (hasShaftWallMarkers)
                {
                    string selectedShaftWallType = GetPresetShaftWallType(effectivePreset);
                    if (string.IsNullOrWhiteSpace(selectedShaftWallType) ||
                        string.Equals(selectedShaftWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                    {
                        notFilled.Add("Стены (Шахта)");
                    }
                    else
                    {
                        WallType matchedShaftWallType = FindWallTypeByName(doc, selectedShaftWallType);
                        if (matchedShaftWallType == null)
                        {
                            otherProblems.Add("Выбран тип стен шахты '" + selectedShaftWallType + "', но такой тип не найден в проекте.");
                        }
                    }
                }

                bool hasLoggiaWallMarkers = HasLoggiaWallMarkersForPlan(doc, targetPlan);
                if (hasLoggiaWallMarkers)
                {
                    string selectedLoggiaWallType = GetPresetLoggiaWallType(effectivePreset);
                    if (string.IsNullOrWhiteSpace(selectedLoggiaWallType) ||
                        string.Equals(selectedLoggiaWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                    {
                        notFilled.Add("Стены (Лоджия)");
                    }
                    else
                    {
                        WallType matchedLoggiaWallType = FindWallTypeByName(doc, selectedLoggiaWallType);
                        if (matchedLoggiaWallType == null)
                        {
                            otherProblems.Add("Выбран тип стен лоджии '" + selectedLoggiaWallType + "', но такой тип не найден в проекте.");
                        }
                    }
                }

                if (doorRequirements.Count > 0)
                {
                    foreach (ApartmentDoorRequirementOption req in doorRequirements
                        .Where(x => x != null)
                        .OrderBy(x => !x.IsEntranceDoor)
                        .ThenBy(x => x.RoomCategory)
                        .ThenBy(x => x.WidthMm)
                        .ThenBy(x => x.DoorTypeName2D))
                    {
                        string selectedDoorTypeName = null;

                        if (effectivePreset.DoorsByRoomCategory != null)
                            effectivePreset.DoorsByRoomCategory.TryGetValue(req.Key, out selectedDoorTypeName);

                        if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                            string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        {
                            notFilled.Add(req.DisplayLabel);
                            continue;
                        }

                        FamilySymbol matchedDoorSymbol = FindDoorSymbolByDisplayNameAndWidth(doc, selectedDoorTypeName, req.WidthMm, req.IsEntranceDoor);
                        if (matchedDoorSymbol == null)
                        {
                            otherProblems.Add(
                                "Для категории '" + req.RoomCategory +
                                "', ширины " + req.WidthMm +
                                " мм выбран тип двери '" + selectedDoorTypeName +
                                "', но такой тип не найден в проекте.");
                        }
                    }
                }
            }

            if (notFilled.Count > 0 || otherProblems.Count > 0)
            {
                List<string> parts = new List<string>();
                parts.Add("Невозможно выполнить построение 3D.");

                if (notFilled.Count > 0)
                {
                    parts.Add("");
                    parts.Add("Не заполнено:");
                    parts.Add("- " + string.Join("\n- ", notFilled.Distinct().ToList()));
                }

                if (otherProblems.Count > 0)
                {
                    parts.Add("");
                    parts.Add("Ошибки:");
                    parts.Add("- " + string.Join("\n- ", otherProblems.Distinct().ToList()));
                }

                validationMessage = string.Join("\n", parts);
                return false;
            }

            return true;
        }

        private void ExecuteConvertTo3D(UIDocument uidoc, Document doc, ApartmentPresetData preset)
        {
            ApartmentPresetData effectivePreset = preset ?? new ApartmentPresetData
            {
                SelectedPlanName = "",
                SelectedPlanModelSignature = "",
                LowerConstraint = "",
                UpperConstraint = "Неприсоединённая",
                BaseOffset = 0,
                WallHeight = 3000,
                WallTypeByThickness = new Dictionary<int, string>(),
                WindowType = "Не выбрано",
                WindowSillHeight = 900,
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = new Dictionary<string, string>(),
                FamilyPostProcessAction = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
            if (targetPlan == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось определить план для построения.");
                return;
            }

            if (targetPlan.GenLevel == null)
            {
                TaskDialog.Show("Ошибка", "У выбранного плана не определён уровень.");
                return;
            }

            Level baseLevel = ResolveBaseLevelForPreset(doc, effectivePreset, targetPlan);
            Level topLevel = ResolveTopLevelForPreset(doc, effectivePreset);

            if (baseLevel == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось определить зависимость снизу.");
                return;
            }

            List<FamilyInstance> apartmentInstances = GetPlacedApartmentInstancesForPlan(doc, targetPlan);
            if (apartmentInstances.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "На выбранном плане не найдено ранее размещённых экземпляров квартир.");
                return;
            }

            List<string> debugMessages = new List<string>();
            List<PreparedApartmentWalls> preparedApartments = new List<PreparedApartmentWalls>();
            List<PreparedApartmentDoors> preparedDoorsByApartment = new List<PreparedApartmentDoors>();
            List<PreparedApartmentWindows> preparedWindowsByApartment = new List<PreparedApartmentWindows>();
            List<PreparedApartmentRooms> preparedRoomsByApartment = new List<PreparedApartmentRooms>();
            Dictionary<long, ApartmentProcessState> apartmentStates = new Dictionary<long, ApartmentProcessState>();

            double connectTol = IDHelper.ConvertMmToInternal(150);
            double intersectionTol = IDHelper.ConvertMmToInternal(10);
            double placementPointZ = 0.0;

            List<ExistingWallLineInfo> existingWalls = GetExistingWallLinesOnLevel(doc, targetPlan.GenLevel.Id);

            foreach (FamilyInstance apartmentFi in apartmentInstances)
            {
                if (apartmentFi == null)
                    continue;

                ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentFi.Id);
                state.ApartmentFamilyName = GetApartmentFamilyReportName(apartmentFi);
                if (state.Restore2DInfo == null)
                    state.Restore2DInfo = BuildRestore2DInfo(apartmentFi, targetPlan);

                try
                {
                    double apartmentWallThicknessInternal = GetApartmentWallThickness(apartmentFi);
                    int apartmentWallThicknessMm = (int)Math.Round(IDHelper.ConvertInternalToMm(apartmentWallThicknessInternal));

                    if (apartmentWallThicknessMm <= 0)
                    {
                        debugMessages.Add("У квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " параметр 'Стены_Толщина' имеет некорректное значение.");
                        continue;
                    }

                    string selectedWallTypeName = GetSelectedWallTypeNameForThickness(effectivePreset, apartmentWallThicknessMm);
                    WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallTypeName, apartmentWallThicknessMm);

                    if (matchedWallType == null)
                    {
                        debugMessages.Add("Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " не найден тип стены '" + selectedWallTypeName + "' с толщиной " + apartmentWallThicknessMm + " мм.");
                        continue;
                    }

                    List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
                    if (roomInstances.Count == 0)
                    {
                        debugMessages.Add("Не найдены вложенные экземпляры 'Помещение' у экземпляра ID = " + IDHelper.ElIdValue(apartmentFi.Id));
                        continue;
                    }

                    List<Line> apartmentAxisLines = new List<Line>();

                    foreach (FamilyInstance roomFi in roomInstances)
                    {
                        try
                        {
                            int skippedWallsForApartment = state.SkippedWallsCount;

                            List<Line> roomAxisLines = BuildPreparedWallAxisLinesForSingleRoom(
                                roomFi,
                                apartmentWallThicknessInternal,
                                existingWalls,
                                targetPlan,
                                connectTol,
                                intersectionTol,
                                debugMessages,
                                ref skippedWallsForApartment);

                            state.SkippedWallsCount = skippedWallsForApartment;

                            if (roomAxisLines == null || roomAxisLines.Count == 0)
                            {
                                state.SkippedRoomsCount++;
                                continue;
                            }

                            apartmentAxisLines.AddRange(roomAxisLines);
                        }
                        catch (Exception exRoom)
                        {
                            state.SkippedRoomsCount++;
                            debugMessages.Add("Ошибка обработки вложенного помещения ID = " + IDHelper.ElIdValue(roomFi.Id) + ": " + exRoom.Message);
                        }
                    }

                    apartmentAxisLines = MergeCollinearLines(apartmentAxisLines);
                    apartmentAxisLines = RemoveSegmentsOverlappingExistingWalls(apartmentAxisLines, existingWalls, apartmentWallThicknessInternal);
                    apartmentAxisLines = MergeCollinearLines(apartmentAxisLines);
                    apartmentAxisLines = WithZ(apartmentAxisLines, placementPointZ);

                    List<FamilyRoomSeparatorMarker> roomSeparatorMarkers = CollectRoomSeparatorMarkersFromApartmentInstance(doc, apartmentFi);
                    List<Line> roomSeparatorLines = BuildRoomSeparatorLinesFromMarkers(roomSeparatorMarkers);
                    roomSeparatorLines = WithZ(roomSeparatorLines, placementPointZ);
                    state.FoundRoomSeparatorsCount = roomSeparatorLines.Count;

                    if (roomSeparatorLines.Count > 0)
                    {
                        apartmentAxisLines = RemoveWallAxisSegmentsAtRoomSeparators(
                            apartmentAxisLines,
                            roomSeparatorLines,
                            apartmentWallThicknessInternal);
                        apartmentAxisLines = MergeCollinearLines(apartmentAxisLines);
                        roomSeparatorLines = TrimRoomSeparatorLinesToWallBoundaries(
                            roomSeparatorLines,
                            apartmentAxisLines,
                            existingWalls,
                            apartmentWallThicknessInternal);
                    }

                    WallType shaftWallType = null;
                    List<Line> shaftAxisLines = new List<Line>();
                    List<FamilyShaftWallMarker> shaftMarkers = CollectShaftWallMarkersFromApartmentInstance(doc, apartmentFi);

                    if (shaftMarkers.Count > 0)
                    {
                        string selectedShaftWallType = GetPresetShaftWallType(effectivePreset);
                        shaftWallType = FindWallTypeByName(doc, selectedShaftWallType);
                        if (shaftWallType == null)
                        {
                            AddApartmentDiagnostic(
                                state,
                                debugMessages,
                                "Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) +
                                " не найден тип стен шахты '" + selectedShaftWallType + "'.");
                        }
                        else
                        {
                            ShaftWallAxisSplitResult shaftSplit = SplitApartmentWallAxesByShaftMarkers(
                                apartmentAxisLines,
                                shaftMarkers,
                                apartmentWallThicknessInternal,
                                shaftWallType);

                            List<Line> shaftReferenceAxisLines = BuildShaftReferenceAxisLines(apartmentAxisLines, existingWalls);
                            apartmentAxisLines = shaftSplit.RegularAxisLines;

                            int skippedShaftMarkers;
                            shaftAxisLines = BuildShaftWallAxesFromMarkers(
                                shaftMarkers,
                                shaftWallType,
                                shaftReferenceAxisLines,
                                out skippedShaftMarkers);
                            shaftAxisLines = MergeCollinearLines(shaftAxisLines);
                            shaftAxisLines = WithZ(shaftAxisLines, placementPointZ);

                            if (skippedShaftMarkers > 0)
                            {
                                AddApartmentDiagnostic(
                                    state,
                                    debugMessages,
                                    "Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) +
                                    " пропущено маркеров шахты без однозначной стороны смещения: " + skippedShaftMarkers + ".");
                            }

                            if (shaftAxisLines.Count == 0)
                            {
                                AddApartmentDiagnostic(
                                    state,
                                    debugMessages,
                                    "Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) +
                                    " найдены маркеры шахты, но не удалось подготовить ни одного участка стены шахты.");
                            }
                        }
                    }

                    WallType loggiaWallType = null;
                    List<Line> loggiaAxisLines = new List<Line>();
                    List<FamilyLoggiaWallMarker> loggiaMarkers = CollectLoggiaWallMarkersFromApartmentInstance(doc, apartmentFi);

                    if (loggiaMarkers.Count > 0)
                    {
                        string selectedLoggiaWallType = GetPresetLoggiaWallType(effectivePreset);
                        loggiaWallType = FindWallTypeByName(doc, selectedLoggiaWallType);
                        if (loggiaWallType == null)
                        {
                            AddApartmentDiagnostic(
                                state,
                                debugMessages,
                                "Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) +
                                " не найден тип стен лоджии '" + selectedLoggiaWallType + "'.");
                        }
                        else
                        {
                            foreach (FamilyLoggiaWallMarker marker in loggiaMarkers)
                            {
                                if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                                    continue;

                                if (Distance2D(marker.ProjectP0, marker.ProjectP1) < IDHelper.ConvertMmToInternal(10))
                                    continue;

                                Line axisLine = Line.CreateBound(marker.ProjectP0, marker.ProjectP1);
                                if (axisLine != null && axisLine.Length > 1e-6)
                                    loggiaAxisLines.Add(axisLine);
                            }

                            loggiaAxisLines = MergeCollinearLines(loggiaAxisLines);
                            loggiaAxisLines = WithZ(loggiaAxisLines, placementPointZ);
                        }
                    }

                    apartmentAxisLines = WithZ(apartmentAxisLines, placementPointZ);
                    shaftAxisLines = WithZ(shaftAxisLines, placementPointZ);
                    loggiaAxisLines = WithZ(loggiaAxisLines, placementPointZ);
                    roomSeparatorLines = WithZ(roomSeparatorLines, placementPointZ);

                    if (apartmentAxisLines.Count == 0 && shaftAxisLines.Count == 0 && roomSeparatorLines.Count == 0)
                    {
                        if (state.SkippedRoomsCount == 0)
                            state.SkippedRoomsCount = roomInstances.Count;

                        debugMessages.Add("Для квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " после покомнатной обработки не осталось осей стен.");
                    }
                    else
                    {
                        preparedApartments.Add(new PreparedApartmentWalls
                        {
                            ApartmentId = apartmentFi.Id,
                            WallType = matchedWallType,
                            ThicknessMm = apartmentWallThicknessMm,
                            AxisLines = apartmentAxisLines,
                            ShaftWallType = shaftWallType,
                            ShaftAxisLines = shaftAxisLines,
                            LoggiaWallType = loggiaWallType,
                            LoggiaAxisLines = loggiaAxisLines,
                            RoomSeparatorLines = roomSeparatorLines
                        });

                        state.HasPreparedWalls = true;
                    }

                    PreparedApartmentDoors preparedDoors = PrepareDoorsForApartment(doc, apartmentFi, effectivePreset, placementPointZ, debugMessages, state);
                    preparedDoorsByApartment.Add(preparedDoors);

                    PreparedApartmentWindows preparedWindows = PrepareWindowsForApartment(doc, apartmentFi, effectivePreset, placementPointZ, debugMessages, state);
                    preparedWindowsByApartment.Add(preparedWindows);

                    PreparedApartmentRooms preparedRooms = PrepareRoomsForApartment(
                        doc,
                        apartmentFi,
                        loggiaAxisLines,
                        apartmentAxisLines,
                        shaftAxisLines,
                        roomSeparatorLines,
                        targetPlan,
                        debugMessages);
                    preparedRoomsByApartment.Add(preparedRooms);
                }
                catch (Exception exApartment)
                {
                    debugMessages.Add("Ошибка обработки квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + ": " + exApartment.Message);
                }
            }

            double baseOffsetInternal = IDHelper.ConvertMmToInternal(effectivePreset.BaseOffset);
            double wallHeightInternal = IDHelper.ConvertMmToInternal(effectivePreset.WallHeight > 0 ? effectivePreset.WallHeight : 3000);

            int totalDoorsPlanned = preparedDoorsByApartment
                .Where(x => x != null && x.Doors != null)
                .Sum(x => x.Doors.Count);

            int installedDoorsCount = 0;

            int totalWindowsPlanned = preparedWindowsByApartment
                .Where(x => x != null && x.Windows != null)
                .Sum(x => x.Windows.Count);

            int installedWindowsCount = 0;

            int totalRoomsPlanned = preparedRoomsByApartment
                .Where(x => x != null && x.Rooms != null)
                .Sum(x => x.Rooms.Count);

            int createdRoomsCount = 0;
            List<RoomAreaMismatchInfo> roomAreaMismatches = new List<RoomAreaMismatchInfo>();

            if (preparedApartments.Count > 0)
            {
                CreatedApartmentWallGeometry createdWallGeometry = CreateWallGeometryInTransaction(
                    doc,
                    preparedApartments,
                    baseLevel,
                    topLevel,
                    baseOffsetInternal,
                    wallHeightInternal,
                    existingWalls,
                    connectTol,
                    apartmentStates);

                createdRoomsCount = PlaceRoomGeometryInTransaction(
                    doc,
                    preparedApartments,
                    preparedRoomsByApartment,
                    targetPlan.GenLevel,
                    roomAreaMismatches,
                    apartmentStates,
                    targetPlan,
                    debugMessages);

                installedDoorsCount = PlaceDoorGeometryInTransaction(
                    doc,
                    preparedApartments,
                    preparedDoorsByApartment,
                    createdWallGeometry.DoorHostWallsByApartment,
                    existingWalls,
                    baseLevel,
                    debugMessages,
                    apartmentStates);

                installedWindowsCount = PlaceWindowGeometryInTransaction(
                    doc,
                    preparedApartments,
                    preparedWindowsByApartment,
                    createdWallGeometry.DoorHostWallsByApartment,
                    existingWalls,
                    baseLevel,
                    debugMessages,
                    apartmentStates);
            }

            ApplyApartmentPostProcessAction(doc, apartmentInstances, effectivePreset.FamilyPostProcessAction, baseLevel, debugMessages, apartmentStates);
            roomAreaMismatches = FilterRoomAreaMismatchesForRoomSeparators(roomAreaMismatches, apartmentStates);
            RefreshRoomAreaMismatchFlags(apartmentStates, roomAreaMismatches);

            List<ApartmentExecutionReportItem> reportItems = BuildExecutionReportItems(doc, apartmentStates, roomAreaMismatches);

            int processedApartmentsCount = apartmentStates.Count;
            int foundEntranceDoorsCount = apartmentStates.Values
                .Where(x => x != null)
                .Sum(x => x.FoundEntranceDoorsCount);
            int installedEntranceDoorsCount = apartmentStates.Values
                .Where(x => x != null)
                .Sum(x => x.InstalledEntranceDoorsCount);
            int foundRoomSeparatorsCount = apartmentStates.Values
                .Where(x => x != null)
                .Sum(x => x.FoundRoomSeparatorsCount);
            int createdRoomSeparatorsCount = apartmentStates.Values
                .Where(x => x != null)
                .Sum(x => x.CreatedRoomSeparatorsCount);
            int roomAreaMismatchCount = roomAreaMismatches.Count;

            ShowExecutionReportWindow(uidoc, targetPlan.Name, processedApartmentsCount, apartmentInstances.Count, createdRoomsCount, totalRoomsPlanned,
                createdRoomSeparatorsCount, foundRoomSeparatorsCount, roomAreaMismatchCount,
                installedDoorsCount, totalDoorsPlanned, foundEntranceDoorsCount, installedEntranceDoorsCount,
                installedWindowsCount, totalWindowsPlanned, reportItems);
        }

        private static List<RoomAreaMismatchInfo> FilterRoomAreaMismatchesForRoomSeparators(
            List<RoomAreaMismatchInfo> roomAreaMismatches,
            Dictionary<long, ApartmentProcessState> apartmentStates)
        {
            if (roomAreaMismatches == null || roomAreaMismatches.Count == 0 || apartmentStates == null || apartmentStates.Count == 0)
                return roomAreaMismatches ?? new List<RoomAreaMismatchInfo>();

            double areaToleranceSquareMeters = 0.1;
            List<RoomAreaMismatchInfo> result = new List<RoomAreaMismatchInfo>();
            List<RoomAreaMismatchInfo> mismatchesWithoutApartment = new List<RoomAreaMismatchInfo>();

            foreach (IGrouping<long, RoomAreaMismatchInfo> group in roomAreaMismatches
                .Where(x => x != null && x.ApartmentId != null)
                .GroupBy(x => IDHelper.ElIdValue(x.ApartmentId)))
            {
                ApartmentProcessState state;
                List<RoomAreaMismatchInfo> mismatchesForApartment = group.ToList();
                bool hasRoomSeparators =
                    apartmentStates.TryGetValue(group.Key, out state) &&
                    state != null &&
                    state.FoundRoomSeparatorsCount > 0;

                if (hasRoomSeparators)
                    mismatchesForApartment = FilterCombinedRoomSeparatorAreaMismatches(mismatchesForApartment, areaToleranceSquareMeters);

                result.AddRange(mismatchesForApartment);
            }

            mismatchesWithoutApartment.AddRange(roomAreaMismatches.Where(x => x == null || x.ApartmentId == null));
            result.AddRange(mismatchesWithoutApartment);
            return result;
        }

        private static List<RoomAreaMismatchInfo> FilterCombinedRoomSeparatorAreaMismatches(List<RoomAreaMismatchInfo> mismatches, double areaToleranceSquareMeters)
        {
            if (mismatches == null || mismatches.Count < 2)
                return mismatches ?? new List<RoomAreaMismatchInfo>();

            HashSet<RoomAreaMismatchInfo> ignoredMismatches = new HashSet<RoomAreaMismatchInfo>();
            List<RoomAreaMismatchInfo> zeroActualMismatches = mismatches
                .Where(x =>
                    x != null &&
                    IDHelper.ConvertInternalAreaToSquareMeters(x.ActualAreaInternal) <= areaToleranceSquareMeters &&
                    IDHelper.ConvertInternalAreaToSquareMeters(x.ExpectedAreaInternal) > areaToleranceSquareMeters)
                .OrderByDescending(x => IDHelper.ConvertInternalAreaToSquareMeters(x.ExpectedAreaInternal))
                .ToList();

            List<RoomAreaMismatchInfo> combinedActualMismatches = mismatches
                .Where(x =>
                    x != null &&
                    IDHelper.ConvertInternalAreaToSquareMeters(x.ActualAreaInternal) - IDHelper.ConvertInternalAreaToSquareMeters(x.ExpectedAreaInternal) > areaToleranceSquareMeters)
                .OrderByDescending(x => IDHelper.ConvertInternalAreaToSquareMeters(x.ActualAreaInternal) - IDHelper.ConvertInternalAreaToSquareMeters(x.ExpectedAreaInternal))
                .ToList();

            foreach (RoomAreaMismatchInfo combinedMismatch in combinedActualMismatches)
            {
                if (ignoredMismatches.Contains(combinedMismatch))
                    continue;

                double expectedArea = IDHelper.ConvertInternalAreaToSquareMeters(combinedMismatch.ExpectedAreaInternal);
                double actualArea = IDHelper.ConvertInternalAreaToSquareMeters(combinedMismatch.ActualAreaInternal);
                double missingPartArea = actualArea - expectedArea;

                if (missingPartArea <= areaToleranceSquareMeters)
                    continue;

                List<RoomAreaMismatchInfo> availableZeroActualMismatches = zeroActualMismatches
                    .Where(x => !ignoredMismatches.Contains(x))
                    .ToList();

                List<RoomAreaMismatchInfo> matchingParts;
                if (!TryFindExpectedAreaSubset(availableZeroActualMismatches, missingPartArea, areaToleranceSquareMeters, out matchingParts))
                    continue;

                ignoredMismatches.Add(combinedMismatch);
                foreach (RoomAreaMismatchInfo matchingPart in matchingParts)
                    ignoredMismatches.Add(matchingPart);
            }

            if (ignoredMismatches.Count == 0)
                return mismatches;

            return mismatches
                .Where(x => x != null && !ignoredMismatches.Contains(x))
                .ToList();
        }

        private static bool TryFindExpectedAreaSubset(
            List<RoomAreaMismatchInfo> candidates,
            double targetAreaSquareMeters,
            double areaToleranceSquareMeters,
            out List<RoomAreaMismatchInfo> result)
        {
            result = new List<RoomAreaMismatchInfo>();

            if (candidates == null || candidates.Count == 0)
                return false;

            return TryFindExpectedAreaSubset(candidates, 0, targetAreaSquareMeters, areaToleranceSquareMeters, new List<RoomAreaMismatchInfo>(), result);
        }

        private static bool TryFindExpectedAreaSubset(
            List<RoomAreaMismatchInfo> candidates,
            int startIndex,
            double remainingAreaSquareMeters,
            double areaToleranceSquareMeters,
            List<RoomAreaMismatchInfo> current,
            List<RoomAreaMismatchInfo> result)
        {
            if (Math.Abs(remainingAreaSquareMeters) <= areaToleranceSquareMeters && current.Count > 0)
            {
                result.AddRange(current);
                return true;
            }

            if (remainingAreaSquareMeters < -areaToleranceSquareMeters)
                return false;

            for (int i = startIndex; i < candidates.Count; i++)
            {
                RoomAreaMismatchInfo candidate = candidates[i];
                if (candidate == null)
                    continue;

                double candidateArea = IDHelper.ConvertInternalAreaToSquareMeters(candidate.ExpectedAreaInternal);
                current.Add(candidate);

                if (TryFindExpectedAreaSubset(candidates, i + 1, remainingAreaSquareMeters - candidateArea, areaToleranceSquareMeters, current, result))
                    return true;

                current.RemoveAt(current.Count - 1);
            }

            return false;
        }

        private static void RefreshRoomAreaMismatchFlags(
            Dictionary<long, ApartmentProcessState> apartmentStates,
            List<RoomAreaMismatchInfo> roomAreaMismatches)
        {
            if (apartmentStates == null || apartmentStates.Count == 0)
                return;

            HashSet<long> apartmentsWithMismatches = new HashSet<long>(
                (roomAreaMismatches ?? new List<RoomAreaMismatchInfo>())
                .Where(x => x != null && x.ApartmentId != null)
                .Select(x => IDHelper.ElIdValue(x.ApartmentId)));

            foreach (ApartmentProcessState state in apartmentStates.Values)
            {
                if (state == null || state.ApartmentId == null)
                    continue;

                state.HasRoomAreaMismatch = apartmentsWithMismatches.Contains(IDHelper.ElIdValue(state.ApartmentId));
            }
        }

        private void ApplyApartmentPostProcessAction(Document doc, List<FamilyInstance> apartmentInstances, ApartmentFamilyPostProcessAction action, Level placementLevel,
            List<string> debugMessages, Dictionary<long, ApartmentProcessState> apartmentStates)
        {
            if (doc == null || apartmentInstances == null || apartmentInstances.Count == 0)
                return;

            if (action == ApartmentFamilyPostProcessAction.Keep2DUnderlay)
                return;

            string transactionName =
                action == ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
                    ? "KPLN. Менеджер квартир. Сохранение 2D-семейств с подложки"
                    : "KPLN. Менеджер квартир. Полное удаление 2D-подложки";

            using (Transaction t = new Transaction(doc, transactionName))
            {
                t.Start();

                foreach (FamilyInstance apartmentFi in apartmentInstances)
                {
                    if (apartmentFi == null)
                        continue;

                    ElementId apartmentId = apartmentFi.Id;
                    if (apartmentId == ElementId.InvalidElementId)
                        continue;

                    if (action == ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay)
                    {
                        ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentId);
                        List<string> furnitureErrors = new List<string>();
                        List<ElementId> createdFurnitureIds = new List<ElementId>();

                        try
                        {
                            CopyFurnitureAndPlumbingFromApartmentUnderlay(doc, apartmentFi, placementLevel, furnitureErrors, createdFurnitureIds);
                        }
                        catch (Exception ex)
                        {
                            furnitureErrors.Add(
                                "Не удалось сохранить 2D-семейства из подложки квартиры ID = " +
                                IDHelper.ElIdValue(apartmentId) + ": " + ex.Message);
                        }

                        if (furnitureErrors.Count > 0)
                        {
                            state.FurnitureErrors.AddRange(furnitureErrors);

                            if (debugMessages != null)
                                debugMessages.AddRange(furnitureErrors);
                        }

                        foreach (ElementId createdFurnitureId in createdFurnitureIds)
                            AddCreatedElementCandidate(state, createdFurnitureId);

                        TryDeleteSource2DApartmentInstance(doc, apartmentId, debugMessages);
                    }
                    else if (action == ApartmentFamilyPostProcessAction.Delete2DUnderlay)
                    {
                        TryDeleteSource2DApartmentInstance(doc, apartmentId, debugMessages);
                    }
                }

                t.Commit();
            }
        }

        private void ShowExecutionReportWindow(UIDocument uidoc, string planName, int processedApartments, int totalApartments, int createdRoomsCount,
            int totalRoomsPlanned, int createdRoomSeparatorsCount, int foundRoomSeparatorsCount,
            int roomAreaMismatchCount,
            int installedDoorsCount, int totalDoorsPlanned, int foundEntranceDoorsCount, int installedEntranceDoorsCount,
            int installedWindowsCount, int totalWindowsPlanned,
            List<ApartmentExecutionReportItem> reportItems)
        {
            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                List<ApartmentExecutionReportItem> items = reportItems != null
                    ? reportItems.ToList()
                    : new List<ApartmentExecutionReportItem>();

                ApartmentExecutionReportItem summaryItem = new ApartmentExecutionReportItem
                {
                    ApartmentId = -1,
                    CustomHeaderText = "Вид: " + (string.IsNullOrWhiteSpace(planName) ? "<без имени>" : planName)
                };

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Обработано квартир: " + processedApartments + " из " + totalApartments
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Создано помещений: " + createdRoomsCount + " из " + totalRoomsPlanned
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Разделитель: " + createdRoomSeparatorsCount + " из " + foundRoomSeparatorsCount
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Несовпадений площадей: " + roomAreaMismatchCount
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Создано дверей: " + installedDoorsCount + " из " + totalDoorsPlanned
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Входных дверей: " + installedEntranceDoorsCount + " из " + foundEntranceDoorsCount
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Создано окон: " + installedWindowsCount + " из " + totalWindowsPlanned
                });

                items.Insert(0, summaryItem);

                ApartmentExecutionReportWindow wnd = new ApartmentExecutionReportWindow(items, new ApartmentExecutionReportActionController());
                wnd.Owner = _window;
                wnd.Show();
                wnd.Activate();
            }));
        }

        private static List<ApartmentExecutionReportItem> BuildExecutionReportItems(
            Document doc,
            Dictionary<long, ApartmentProcessState> apartmentStates,
            List<RoomAreaMismatchInfo> roomAreaMismatches)
        {
            List<ApartmentExecutionReportItem> result = new List<ApartmentExecutionReportItem>();

            if (apartmentStates == null || apartmentStates.Count == 0)
                return result;

            foreach (ApartmentProcessState state in apartmentStates.Values.OrderBy(x => IDHelper.ElIdValue(x.ApartmentId)))
            {
                if (state == null || state.ApartmentId == null || state.ApartmentId == ElementId.InvalidElementId)
                    continue;

                long apartmentId = IDHelper.ElIdValue(state.ApartmentId);

                ApartmentExecutionReportItem reportItem = new ApartmentExecutionReportItem
                {
                    ApartmentId = apartmentId,
                    CustomHeaderText = !string.IsNullOrWhiteSpace(state.ApartmentFamilyName)
                        ? state.ApartmentFamilyName
                        : "Квартира [" + apartmentId + "]"
                };

                AddExistingNavigationCandidatesToReport(doc, reportItem, state);
                AddExistingDeletableCandidatesToReport(doc, reportItem, state);

                if (state.Restore2DInfo != null && doc != null && doc.GetElement(state.ApartmentId) == null)
                    reportItem.Restore2DInfo = state.Restore2DInfo;

                bool hasSkippedRooms = state.SkippedRoomsCount > 0;
                bool hasSkippedDoors = state.SkippedDoorsCount > 0;
                bool hasSkippedWindows = state.SkippedWindowsCount > 0;
                bool hasFurnitureErrors = state.FurnitureErrors != null && state.FurnitureErrors.Count > 0;
                bool hasErrorMessages = state.ErrorMessages != null && state.ErrorMessages.Count > 0;
                bool hasAreaMismatches = state.HasRoomAreaMismatch;
                bool hasReportableErrors = hasSkippedRooms || hasSkippedDoors || hasSkippedWindows || hasFurnitureErrors || hasErrorMessages || hasAreaMismatches;

                if (hasReportableErrors)
                {
                    if (hasSkippedRooms)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не построено помещений: " + state.SkippedRoomsCount,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });
                    }

                    if (hasSkippedDoors)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не построено дверей: " + state.SkippedDoorsCount,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });
                    }

                    if (hasSkippedWindows)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не построено окон: " + state.SkippedWindowsCount,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });
                    }

                    if (hasFurnitureErrors)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не создано мебели/сантехники: " + state.FurnitureErrors.Count,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });

                        foreach (string furnitureError in state.FurnitureErrors
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct()
                            .Take(10))
                        {
                            reportItem.Lines.Add(new ApartmentExecutionReportLine
                            {
                                Text = "-- " + furnitureError,
                                Foreground = System.Windows.Media.Brushes.DarkOrange
                            });
                        }
                    }

                    if (hasErrorMessages)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Диагностика:",
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });

                        foreach (string errorMessage in state.ErrorMessages
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct()
                            .Take(20))
                        {
                            reportItem.Lines.Add(new ApartmentExecutionReportLine
                            {
                                Text = "-- " + errorMessage,
                                Foreground = System.Windows.Media.Brushes.DarkOrange
                            });
                        }
                    }
                }

                if (roomAreaMismatches != null)
                {
                    List<RoomAreaMismatchInfo> mismatchesForApartment = roomAreaMismatches
                        .Where(x =>
                            x != null &&
                            x.ApartmentId != null &&
                            IDHelper.ElIdValue(x.ApartmentId) == reportItem.ApartmentId)
                        .OrderBy(x => x.RoomName)
                        .ToList();

                    if (mismatchesForApartment.Count > 0)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Несовпадения площадей: " + mismatchesForApartment.Count,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });

                    }

                    foreach (RoomAreaMismatchInfo mismatchItem in mismatchesForApartment)
                    {
                        double expectedArea = IDHelper.ConvertInternalAreaToSquareMeters(mismatchItem.ExpectedAreaInternal);
                        double actualArea = IDHelper.ConvertInternalAreaToSquareMeters(mismatchItem.ActualAreaInternal);
                        string expectedText = expectedArea.ToString("0.##");
                        string actualText = actualArea.ToString("0.##");
                        string differenceText = (actualArea - expectedArea).ToString("+0.##;-0.##;0");

                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "-- Площадь: " +
                                   (string.IsNullOrWhiteSpace(mismatchItem.RoomName) ? "Помещение" : mismatchItem.RoomName) +
                                   " | 2D = " + expectedText + " м², 3D = " + actualText + " м², Δ = " + differenceText + " м²",
                            Foreground = System.Windows.Media.Brushes.DarkOrange
                        });
                    }
                }

                if (reportItem.Lines.Count > 0)
                    result.Add(reportItem);
            }

            return result;
        }

        private static ApartmentProcessState GetOrCreateApartmentState(Dictionary<long, ApartmentProcessState> states, ElementId apartmentId)
        {
            long key = IDHelper.ElIdValue(apartmentId);
            ApartmentProcessState state;

            if (!states.TryGetValue(key, out state))
            {
                state = new ApartmentProcessState
                {
                    ApartmentId = apartmentId,
                    NavigationElementId = apartmentId
                };

                states.Add(key, state);
            }

            AddNavigationElementCandidate(state, apartmentId);

            return state;
        }

        private static string GetApartmentFamilyReportName(FamilyInstance apartmentFi)
        {
            if (apartmentFi == null)
                return null;

            if (apartmentFi.Symbol != null)
            {
                if (apartmentFi.Symbol.Family != null && !string.IsNullOrWhiteSpace(apartmentFi.Symbol.Family.Name))
                    return apartmentFi.Symbol.Family.Name.Trim();

                if (!string.IsNullOrWhiteSpace(apartmentFi.Symbol.Name))
                    return apartmentFi.Symbol.Name.Trim();
            }

            return null;
        }

        private static void AddApartmentDiagnostic(ApartmentProcessState state, List<string> debugMessages, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            if (debugMessages != null)
                debugMessages.Add(message);

            if (state == null)
                return;

            if (state.ErrorMessages == null)
                state.ErrorMessages = new List<string>();

            if (!state.ErrorMessages.Contains(message))
                state.ErrorMessages.Add(message);
        }

        private static void AddNavigationElementCandidate(ApartmentProcessState state, ElementId elementId)
        {
            if (state == null || elementId == null || elementId == ElementId.InvalidElementId)
                return;

            if (state.NavigationElementId == null || state.NavigationElementId == ElementId.InvalidElementId)
                state.NavigationElementId = elementId;

            if (state.NavigationElementIds == null)
                state.NavigationElementIds = new List<ElementId>();

            long value = IDHelper.ElIdValue(elementId);
            if (state.NavigationElementIds.Any(x => x != null && IDHelper.ElIdValue(x) == value))
                return;

            state.NavigationElementIds.Add(elementId);
        }

        private static void AddCreatedElementCandidate(ApartmentProcessState state, ElementId elementId)
        {
            AddNavigationElementCandidate(state, elementId);

            if (state == null || elementId == null || elementId == ElementId.InvalidElementId)
                return;

            if (state.CreatedElementIds == null)
                state.CreatedElementIds = new List<ElementId>();

            long value = IDHelper.ElIdValue(elementId);
            if (state.CreatedElementIds.Any(x => x != null && IDHelper.ElIdValue(x) == value))
                return;

            state.CreatedElementIds.Add(elementId);
        }

        private static void AddExistingNavigationCandidatesToReport(Document doc, ApartmentExecutionReportItem reportItem, ApartmentProcessState state)
        {
            if (reportItem == null || state == null)
                return;

            List<ElementId> candidates = new List<ElementId>();

            if (state.NavigationElementIds != null)
                candidates.AddRange(state.NavigationElementIds.Where(x => x != null && x != ElementId.InvalidElementId));

            if (state.NavigationElementId != null && state.NavigationElementId != ElementId.InvalidElementId)
                candidates.Add(state.NavigationElementId);

            if (state.ApartmentId != null && state.ApartmentId != ElementId.InvalidElementId)
                candidates.Add(state.ApartmentId);

            foreach (ElementId candidate in candidates)
            {
                if (candidate == null || candidate == ElementId.InvalidElementId)
                    continue;

                if (doc != null && doc.GetElement(candidate) == null)
                    continue;

                long value = IDHelper.ElIdValue(candidate);
                if (reportItem.NavigationElementIds.Contains(value))
                    continue;

                reportItem.NavigationElementIds.Add(value);

                if (reportItem.NavigationElementId <= 0)
                    reportItem.NavigationElementId = value;
            }
        }

        private static void AddExistingDeletableCandidatesToReport(Document doc, ApartmentExecutionReportItem reportItem, ApartmentProcessState state)
        {
            if (doc == null || reportItem == null || state == null || state.CreatedElementIds == null)
                return;

            foreach (ElementId candidate in state.CreatedElementIds)
            {
                if (candidate == null || candidate == ElementId.InvalidElementId)
                    continue;

                if (doc.GetElement(candidate) == null)
                    continue;

                long value = IDHelper.ElIdValue(candidate);
                if (reportItem.DeletableElementIds.Contains(value))
                    continue;

                reportItem.DeletableElementIds.Add(value);
            }
        }

        private static Apartment2DRestoreInfo BuildRestore2DInfo(FamilyInstance apartmentFi, ViewPlan targetPlan)
        {
            if (apartmentFi == null || targetPlan == null || apartmentFi.Symbol == null)
                return null;

            Transform transform = apartmentFi.GetTransform();
            if (transform == null || transform.Origin == null)
                return null;

            XYZ origin = transform.Origin;
            XYZ basisX = transform.BasisX;
            double rotation = basisX != null
                ? Math.Atan2(basisX.Y, basisX.X)
                : 0.0;

            return new Apartment2DRestoreInfo
            {
                SymbolId = IDHelper.ElIdValue(apartmentFi.Symbol.Id),
                ViewId = IDHelper.ElIdValue(targetPlan.Id),
                LevelId = targetPlan.GenLevel != null ? IDHelper.ElIdValue(targetPlan.GenLevel.Id) : 0,
                X = origin.X,
                Y = origin.Y,
                Z = origin.Z,
                Rotation = rotation
            };
        }
    }
}
