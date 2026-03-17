using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows.Interop;

namespace KPLN_Tools.ExecutableCommand
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandApartmentManagerShow : IExternalCommand
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string ApartmentInstanceMarkerPrefix = "[KPLN_APT_INSTANCE_ID=";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                var wnd = new KPLN_Tools.Forms.ApartmentManagerWindow(DBWorkerService.CurrentDBUserSubDepartment.Id);
                new WindowInteropHelper(wnd).Owner = uiapp.MainWindowHandle;

                bool? res = wnd.ShowDialog();
                if (res != true)
                    return Result.Succeeded;

                if (wnd.ConvertTo3DRequested)
                {
                    ViewPlan floorPlan = doc.ActiveView as ViewPlan;
                    if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
                    {
                        TaskDialog.Show("Предупреждение", "Откройте план этажа перед построением стен.");
                        return Result.Cancelled;
                    }

                    if (floorPlan.GenLevel == null)
                    {
                        TaskDialog.Show("Ошибка", "У активного плана не определён уровень.");
                        return Result.Failed;
                    }

                    var preset = wnd.ApartmentPresetData;

                    List<FamilyInstance> apartmentInstances = GetPlacedApartmentInstancesOnLevel(doc, floorPlan.GenLevel);
                    if (apartmentInstances.Count == 0)
                    {
                        TaskDialog.Show("Apartment Manager", "На текущем уровне не найдено ранее размещённых экземпляров квартир.");
                        return Result.Cancelled;
                    }

                    List<string> debugMessages = new List<string>();
                    List<PreparedApartmentWalls> preparedApartments = new List<PreparedApartmentWalls>();

                    double connectTol = ConvertMmToInternal(150);
                    List<ExistingWallLineInfo> existingWalls = GetExistingWallLinesOnLevel(doc, floorPlan.GenLevel.Id);

                    WallType preferredWallType = null;
                    if (preset != null && !string.IsNullOrWhiteSpace(preset.WallType) && preset.WallType != "Не выбрано")
                        preferredWallType = FindWallTypeByName(doc, preset.WallType);

                    int roomLoopsCount = 0;

                    foreach (FamilyInstance apartmentFi in apartmentInstances)
                    {
                        try
                        {
                            double apartmentWallThicknessInternal = GetApartmentWallThickness(apartmentFi);
                            int apartmentWallThicknessMm = (int)Math.Round(ConvertInternalToMm(apartmentWallThicknessInternal));

                            if (apartmentWallThicknessMm <= 0)
                            {
                                debugMessages.Add(
                                    "У квартиры Id=" + GetElementIdValue(apartmentFi.Id) +
                                    " параметр 'Стены_Толщина' имеет некорректное значение.");
                                continue;
                            }

                            WallType matchedWallType = FindWallTypeByThickness(doc, apartmentWallThicknessMm, preferredWallType);
                            if (matchedWallType == null)
                            {
                                debugMessages.Add(
                                    "Для квартиры Id=" + GetElementIdValue(apartmentFi.Id) +
                                    " не найден тип стены с толщиной " + apartmentWallThicknessMm + " мм.");
                                continue;
                            }

                            List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
                            if (roomInstances.Count == 0)
                            {
                                debugMessages.Add(
                                    "Не найдены вложенные экземпляры 'Помещение' у экземпляра Id=" +
                                    GetElementIdValue(apartmentFi.Id));
                                continue;
                            }

                            List<CurveLoop> apartmentRoomLoops = new List<CurveLoop>();

                            foreach (FamilyInstance roomFi in roomInstances)
                            {
                                try
                                {
                                    CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi);
                                    apartmentRoomLoops.Add(roomLoop);
                                    roomLoopsCount++;
                                }
                                catch (Exception exRoom)
                                {
                                    debugMessages.Add(
                                        "Ошибка обработки вложенного помещения Id=" +
                                        GetElementIdValue(roomFi.Id) + ": " + exRoom.Message);
                                }
                            }

                            if (apartmentRoomLoops.Count == 0)
                                continue;

                            List<Line> wallAxisLines = BuildClosedWallAxisLinesFromRooms(
                                apartmentRoomLoops,
                                apartmentWallThicknessInternal,
                                debugMessages);

                            if (wallAxisLines.Count == 0)
                            {
                                debugMessages.Add(
                                    "Для квартиры Id=" + GetElementIdValue(apartmentFi.Id) +
                                    " не удалось вычислить оси стен.");
                                continue;
                            }

                            List<Line> preparedAxisLines = SnapNewLinesToExistingWalls(wallAxisLines, existingWalls, connectTol);
                            preparedAxisLines = MergeCollinearLines(preparedAxisLines);
                            preparedAxisLines = RemoveSegmentsOverlappingExistingWalls(preparedAxisLines, existingWalls);
                            preparedAxisLines = MergeCollinearLines(preparedAxisLines);

                            if (preparedAxisLines.Count == 0)
                            {
                                debugMessages.Add(
                                    "Для квартиры Id=" + GetElementIdValue(apartmentFi.Id) +
                                    " все вычисленные оси перекрыты существующими стенами.");
                                continue;
                            }

                            preparedApartments.Add(new PreparedApartmentWalls
                            {
                                ApartmentId = apartmentFi.Id,
                                WallType = matchedWallType,
                                ThicknessMm = apartmentWallThicknessMm,
                                AxisLines = preparedAxisLines
                            });
                        }
                        catch (Exception exApartment)
                        {
                            debugMessages.Add(
                                "Ошибка обработки квартиры Id=" +
                                GetElementIdValue(apartmentFi.Id) + ": " + exApartment.Message);
                        }
                    }

                    if (preparedApartments.Count == 0)
                    {
                        string debugText = debugMessages.Count > 0
                            ? "\n\n" + string.Join("\n", debugMessages)
                            : "";
                        TaskDialog.Show("Apartment Manager", "Не удалось подготовить ни одной стены." + debugText);
                        return Result.Cancelled;
                    }

                    double wallHeightInternal = ConvertMmToInternal(preset.WallHeight);

                    int createdWalls = 0;
                    int preparedAxisCount = 0;

                    using (Transaction t = new Transaction(doc, "KPLN - Построение стен по помещениям"))
                    {
                        t.Start();

                        foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                        {
                            if (apartmentWalls == null || apartmentWalls.WallType == null || apartmentWalls.AxisLines == null)
                                continue;

                            foreach (Line axis in apartmentWalls.AxisLines)
                            {
                                if (axis == null || axis.Length < 1e-6)
                                    continue;

                                Wall.Create(
                                    doc,
                                    axis,
                                    apartmentWalls.WallType.Id,
                                    floorPlan.GenLevel.Id,
                                    wallHeightInternal,
                                    0,
                                    false,
                                    false);

                                createdWalls++;
                                preparedAxisCount++;
                            }
                        }

                        t.Commit();
                    }

                    string report =
                        "Построение завершено.\n" +
                        "Экземпляров квартир: " + apartmentInstances.Count + "\n" +
                        "Контуров помещений: " + roomLoopsCount + "\n" +
                        "Подготовлено наборов стен: " + preparedApartments.Count + "\n" +
                        "Стеновых осей после подготовки: " + preparedAxisCount + "\n" +
                        "Создано стен: " + createdWalls + "\n" +
                        "Найдено существующих стен на уровне: " + existingWalls.Count;

                    if (debugMessages.Count > 0)
                        report += "\n\nЗамечания:\n" + string.Join("\n", debugMessages);

                    TaskDialog.Show("ConvertTo3D", report);
                    return Result.Succeeded;
                }
                else
                {
                    int id = wnd.SelectedApartmentId;

                    ViewPlan floorPlan = doc.ActiveView as ViewPlan;
                    if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
                    {
                        TaskDialog.Show("Предупреждение", "Откройте план этажа перед размещением семейства.");
                        return Result.Cancelled;
                    }

                    string familyPath = GetFamilyPathById(id);
                    if (string.IsNullOrWhiteSpace(familyPath))
                    {
                        TaskDialog.Show("Ошибка", "Для ID=" + id + " не найден FPATH в базе.");
                        return Result.Failed;
                    }

                    if (!File.Exists(familyPath))
                    {
                        TaskDialog.Show("Ошибка", "Файл семейства не найден:\n" + familyPath);
                        return Result.Failed;
                    }

                    XYZ insertPoint;
                    try
                    {
                        insertPoint = uidoc.Selection.PickPoint("Укажите точку вставки семейства");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }

                    using (Transaction t = new Transaction(doc, "Разместить семейство квартиры"))
                    {
                        t.Start();

                        Family family = LoadOrFindFamily(doc, familyPath);
                        if (family == null)
                            throw new Exception("Не удалось загрузить или найти семейство в проекте.");

                        FamilySymbol symbol = GetFirstFamilySymbol(doc, family);
                        if (symbol == null)
                            throw new Exception("У семейства не найдено ни одного типоразмера.");

                        if (!symbol.IsActive)
                        {
                            symbol.Activate();
                            doc.Regenerate();
                        }

                        FamilyInstance placedInstance = PlaceFamilyInstance(doc, floorPlan, symbol, insertPoint);
                        if (placedInstance == null)
                            throw new Exception("Не удалось разместить экземпляр семейства.");

                        MarkPlacedApartmentInstance(placedInstance, id);

                        t.Commit();
                    }

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private static string GetFamilyPathById(int id)
        {
            if (!File.Exists(DbPath))
                throw new FileNotFoundException("Не найдена база данных", DbPath);

            using (var con = OpenConnection(DbPath, true))
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT FPATH FROM Main WHERE ID = @id LIMIT 1;";
                cmd.Parameters.AddWithValue("@id", id);

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    return null;

                return result.ToString().Trim();
            }
        }

        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? ("Data Source=" + dbPath + ";Version=3;Read Only=True;")
                : ("Data Source=" + dbPath + ";Version=3;");

            SQLiteConnection con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

        private static Family LoadOrFindFamily(Document doc, string familyPath)
        {
            string familyName = Path.GetFileNameWithoutExtension(familyPath);

            Family existingFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => string.Equals(f.Name, familyName, StringComparison.OrdinalIgnoreCase));

            if (existingFamily != null)
                return existingFamily;

            Family loadedFamily;
            bool loaded = doc.LoadFamily(familyPath, out loadedFamily);

            if (loaded && loadedFamily != null)
                return loadedFamily;

            existingFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => string.Equals(f.Name, familyName, StringComparison.OrdinalIgnoreCase));

            return existingFamily;
        }

        private static FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
            if (symbolIds == null || symbolIds.Count == 0)
                return null;

            ElementId firstId = symbolIds.First();
            return doc.GetElement(firstId) as FamilySymbol;
        }

        private static FamilyInstance PlaceFamilyInstance(Document doc, ViewPlan floorPlan, FamilySymbol symbol, XYZ point)
        {
            FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;

            switch (placementType)
            {
                case FamilyPlacementType.ViewBased:
                    return doc.Create.NewFamilyInstance(point, symbol, floorPlan);

                case FamilyPlacementType.OneLevelBased:
                case FamilyPlacementType.OneLevelBasedHosted:
                    if (floorPlan.GenLevel == null)
                        throw new Exception("У активного плана не определён уровень.");

                    return doc.Create.NewFamilyInstance(
                        point,
                        symbol,
                        floorPlan.GenLevel,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                case FamilyPlacementType.WorkPlaneBased:
                    if (floorPlan.GenLevel == null)
                        throw new Exception("У активного плана не определён уровень.");

                    return doc.Create.NewFamilyInstance(
                        point,
                        symbol,
                        floorPlan.GenLevel,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                default:
                    throw new NotSupportedException(
                        "Тип размещения семейства не поддержан: " + placementType +
                        ". Сейчас поддержаны ViewBased, OneLevelBased, OneLevelBasedHosted, WorkPlaneBased.");
            }
        }

        private static void MarkPlacedApartmentInstance(FamilyInstance fi, int apartmentId)
        {
            string marker = ApartmentInstanceMarkerPrefix + apartmentId + "]";
            AppendComment(fi, marker);
        }

        private static void AppendComment(Element e, string textToAppend)
        {
            if (e == null || string.IsNullOrWhiteSpace(textToAppend))
                return;

            Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p == null || p.IsReadOnly)
                return;

            string oldValue = p.AsString();
            if (string.IsNullOrWhiteSpace(oldValue))
            {
                p.Set(textToAppend);
                return;
            }

            if (oldValue.Contains(textToAppend))
                return;

            p.Set(oldValue + " " + textToAppend);
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

                if (!comment.Contains(ApartmentInstanceMarkerPrefix))
                    continue;

                ElementId levelId = GetInstanceLevelId(fi);
                if (levelId == ElementId.InvalidElementId)
                    continue;

                if (levelId == level.Id)
                    result.Add(fi);
            }

            return result;
        }

        private static ElementId GetInstanceLevelId(FamilyInstance fi)
        {
            if (fi == null)
                return ElementId.InvalidElementId;

            Parameter p =
                fi.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                fi.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);

            if (p != null && p.StorageType == StorageType.ElementId)
                return p.AsElementId();

            if (fi.LevelId != ElementId.InvalidElementId)
                return fi.LevelId;

            return ElementId.InvalidElementId;
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

        private static CurveLoop BuildRoomLoopFromInstance(FamilyInstance roomFi)
        {
            if (roomFi == null)
                throw new ArgumentNullException("roomFi");

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                throw new Exception("Не удалось получить Transform для вложенного помещения.");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            XYZ p1 = tr.OfPoint(new XYZ(-halfW, -halfD, 0));
            XYZ p2 = tr.OfPoint(new XYZ(halfW, -halfD, 0));
            XYZ p3 = tr.OfPoint(new XYZ(halfW, halfD, 0));
            XYZ p4 = tr.OfPoint(new XYZ(-halfW, halfD, 0));

            List<Curve> profile = new List<Curve>();
            profile.Add(Line.CreateBound(p1, p2));
            profile.Add(Line.CreateBound(p2, p3));
            profile.Add(Line.CreateBound(p3, p4));
            profile.Add(Line.CreateBound(p4, p1));

            return CurveLoop.Create(profile);
        }

        private static double GetRequiredLengthParam(Element e, params string[] paramNames)
        {
            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            Element typeElem = null;
            if (e != null && e.Document != null)
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
                        return p.AsDouble();
                }
            }

            throw new Exception("Не найден параметр длины: " + string.Join(", ", paramNames));
        }

        private static double GetApartmentWallThickness(FamilyInstance apartmentFi)
        {
            if (apartmentFi == null)
                throw new ArgumentNullException("apartmentFi");

            Parameter p = apartmentFi.LookupParameter("Стены_Толщина");
            if (p != null && p.StorageType == StorageType.Double)
                return p.AsDouble();

            Element typeElem = apartmentFi.Document.GetElement(apartmentFi.GetTypeId());
            if (typeElem != null)
            {
                p = typeElem.LookupParameter("Стены_Толщина");
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            throw new Exception("Не найден параметр 'Стены_Толщина' у экземпляра или типа семейства квартиры.");
        }

        private static WallType FindWallTypeByName(Document doc, string wallTypeName)
        {
            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            if (allWallTypes.Count == 0)
                return null;

            if (string.IsNullOrWhiteSpace(wallTypeName) || wallTypeName == "Не выбрано")
                return allWallTypes.First();

            WallType exact = allWallTypes
                .FirstOrDefault(x => string.Equals(x.Name, wallTypeName, StringComparison.OrdinalIgnoreCase));

            if (exact != null)
                return exact;

            WallType contains = allWallTypes
                .FirstOrDefault(x => x.Name.IndexOf(wallTypeName, StringComparison.OrdinalIgnoreCase) >= 0);

            if (contains != null)
                return contains;

            return null;
        }

        private static WallType FindWallTypeByThickness(Document doc, int thicknessMm, WallType preferredWallType = null)
        {
            if (preferredWallType != null)
            {
                int preferredThickness;
                if (TryGetThicknessFromWallTypeName(preferredWallType.Name, out preferredThickness) &&
                    preferredThickness == thicknessMm)
                {
                    return preferredWallType;
                }
            }

            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            foreach (WallType wt in allWallTypes)
            {
                if (wt == null || string.IsNullOrWhiteSpace(wt.Name))
                    continue;

                int parsedThickness;
                if (!TryGetThicknessFromWallTypeName(wt.Name, out parsedThickness))
                    continue;

                if (parsedThickness == thicknessMm)
                    return wt;
            }

            return null;
        }

        private static bool TryGetThicknessFromWallTypeName(string wallTypeName, out int thicknessMm)
        {
            thicknessMm = 0;

            if (string.IsNullOrWhiteSpace(wallTypeName))
                return false;

            string[] parts = wallTypeName
                .Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToArray();

            if (parts.Length == 0)
                return false;

            for (int i = parts.Length - 1; i >= 0; i--)
            {
                int value;
                if (int.TryParse(parts[i], out value))
                {
                    thicknessMm = value;
                    return true;
                }
            }

            return false;
        }

        private static List<Line> BuildClosedWallAxisLinesFromRooms(List<CurveLoop> roomLoops, double wallWidth, List<string> debugMessages)
        {
            List<Line> result = new List<Line>();
            double halfWidth = wallWidth / 2.0;

            foreach (CurveLoop roomLoop in roomLoops)
            {
                try
                {
                    List<Line> offsetLoopLines = BuildOffsetClosedLoop(roomLoop, halfWidth);
                    result.AddRange(offsetLoopLines);
                }
                catch (Exception ex)
                {
                    debugMessages.Add("Ошибка построения замкнутого контура стены: " + ex.Message);
                }
            }

            return MergeCollinearLines(result);
        }

        private static List<Line> BuildOffsetClosedLoop(CurveLoop loop, double offset)
        {
            const double tol = 1e-9;

            List<Line> edges = loop
                .Cast<Curve>()
                .Select(x => x as Line)
                .Where(x => x != null && x.Length > tol)
                .ToList();

            if (edges.Count < 3)
                throw new Exception("Контур помещения должен содержать минимум 3 линейных сегмента.");

            List<XYZ> vertices = ExtractOrderedVertices(edges);
            if (vertices.Count < 3)
                throw new Exception("Не удалось извлечь вершины контура помещения.");

            bool ccw = GetSignedAreaXY(vertices) > 0.0;

            List<OffsetLine2D> offsetLines = new List<OffsetLine2D>();
            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];

                XYZ dir = Normalize2D(b - a);
                if (dir == null)
                    throw new Exception("Обнаружен нулевой сегмент в контуре.");

                XYZ outward = ccw
                    ? new XYZ(dir.Y, -dir.X, 0)
                    : new XYZ(-dir.Y, dir.X, 0);

                XYZ oa = new XYZ(a.X + outward.X * offset, a.Y + outward.Y * offset, a.Z);
                XYZ ob = new XYZ(b.X + outward.X * offset, b.Y + outward.Y * offset, b.Z);

                offsetLines.Add(new OffsetLine2D
                {
                    P0 = oa,
                    P1 = ob,
                    Dir = dir,
                    Z = a.Z
                });
            }

            List<XYZ> offsetVertices = new List<XYZ>();
            for (int i = 0; i < offsetLines.Count; i++)
            {
                OffsetLine2D prev = offsetLines[(i - 1 + offsetLines.Count) % offsetLines.Count];
                OffsetLine2D cur = offsetLines[i];

                XYZ intersection;
                if (!TryIntersectInfiniteLines2D(prev.P0, prev.Dir, cur.P0, cur.Dir, out intersection))
                    intersection = cur.P0;

                double z = 0.5 * (prev.Z + cur.Z);
                offsetVertices.Add(new XYZ(intersection.X, intersection.Y, z));
            }

            List<Line> result = new List<Line>();
            for (int i = 0; i < offsetVertices.Count; i++)
            {
                XYZ p0 = offsetVertices[i];
                XYZ p1 = offsetVertices[(i + 1) % offsetVertices.Count];

                if (p0.DistanceTo(p1) > tol)
                    result.Add(Line.CreateBound(p0, p1));
            }

            return result;
        }

        private static List<XYZ> ExtractOrderedVertices(List<Line> lines)
        {
            const double tol = 1e-6;
            List<XYZ> pts = new List<XYZ>();

            if (lines == null || lines.Count == 0)
                return pts;

            foreach (Line line in lines)
            {
                XYZ p = line.GetEndPoint(0);
                if (pts.Count == 0 || pts[pts.Count - 1].DistanceTo(p) > tol)
                    pts.Add(p);
            }

            if (pts.Count >= 2 && pts[0].DistanceTo(pts[pts.Count - 1]) < tol)
                pts.RemoveAt(pts.Count - 1);

            return pts;
        }

        private static List<ExistingWallLineInfo> GetExistingWallLinesOnLevel(Document doc, ElementId levelId)
        {
            List<ExistingWallLineInfo> result = new List<ExistingWallLineInfo>();

            IEnumerable<Wall> walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>();

            foreach (Wall wall in walls)
            {
                if (wall == null)
                    continue;

                if (wall.LevelId != levelId)
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line line = lc.Curve as Line;
                if (line == null || line.Length < 1e-9)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                result.Add(new ExistingWallLineInfo
                {
                    WallId = wall.Id,
                    P0 = p0,
                    P1 = p1,
                    Dir = dir,
                    Z = 0.5 * (p0.Z + p1.Z)
                });
            }

            return result;
        }

        private static List<Line> SnapNewLinesToExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            List<Line> result = new List<Line>();

            foreach (Line line in newLines)
            {
                if (line == null || line.Length < 1e-9)
                    continue;

                Line snapped = SnapSingleLineToExistingWalls(line, existingWalls, snapTol);
                if (snapped != null && snapped.Length > 1e-9)
                    result.Add(snapped);
            }

            return result;
        }

        private static Line SnapSingleLineToExistingWalls(Line line, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);
            if (dir == null)
                return line;

            XYZ newP0 = SnapEndpointToExistingWalls(p0, new XYZ(-dir.X, -dir.Y, 0), line, existingWalls, snapTol);
            XYZ newP1 = SnapEndpointToExistingWalls(p1, dir, line, existingWalls, snapTol);

            if (newP0.DistanceTo(newP1) < 1e-9)
                return null;

            return Line.CreateBound(newP0, newP1);
        }

        private static XYZ SnapEndpointToExistingWalls(
            XYZ endpoint,
            XYZ extensionDir,
            Line sourceLine,
            List<ExistingWallLineInfo> existingWalls,
            double snapTol)
        {
            const double tol = 1e-9;

            XYZ bestPoint = endpoint;
            double bestDist = double.MaxValue;

            XYZ sourceP0 = sourceLine.GetEndPoint(0);
            XYZ sourceP1 = sourceLine.GetEndPoint(1);
            XYZ sourceDir = Normalize2D(sourceP1 - sourceP0);
            if (sourceDir == null)
                return endpoint;

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (Math.Abs(ex.Z - endpoint.Z) > snapTol)
                    continue;

                XYZ inter;
                if (TryIntersectInfiniteLines2D(endpoint, extensionDir, ex.P0, ex.Dir, out inter))
                {
                    if (PointOnSegment2D(inter, ex.P0, ex.P1, snapTol))
                    {
                        XYZ delta = inter - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, inter);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(inter.X, inter.Y, endpoint.Z);
                        }
                    }
                }

                if (AreCollinear2D(sourceP0, sourceP1, ex.P0, ex.P1, snapTol))
                {
                    XYZ[] candidates = new XYZ[] { ex.P0, ex.P1 };
                    foreach (XYZ c in candidates)
                    {
                        XYZ delta = c - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, c);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(c.X, c.Y, endpoint.Z);
                        }
                    }
                }
            }

            return bestPoint;
        }

        private static List<Line> RemoveSegmentsOverlappingExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls)
        {
            const double tol = 1e-6;
            List<Line> result = new List<Line>();

            foreach (Line newLine in newLines)
            {
                if (newLine == null || newLine.Length <= tol)
                    continue;

                List<Line> remaining = SubtractExistingWallsFromNewLine(newLine, existingWalls);
                foreach (Line part in remaining)
                {
                    if (part != null && part.Length > tol)
                        result.Add(part);
                }
            }

            return result;
        }

        private static List<Line> SubtractExistingWallsFromNewLine(Line newLine, List<ExistingWallLineInfo> existingWalls)
        {
            const double tol = 1e-6;

            XYZ p0 = newLine.GetEndPoint(0);
            XYZ p1 = newLine.GetEndPoint(1);

            XYZ dir = Normalize2D(p1 - p0);
            if (dir == null)
                return new List<Line>();

            dir = CanonicalizeDirection(dir);
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            double newOffset = Dot2D(p0, normal);
            double newZ = 0.5 * (p0.Z + p1.Z);

            double newT0 = Dot2D(p0, dir);
            double newT1 = Dot2D(p1, dir);
            double newFrom = Math.Min(newT0, newT1);
            double newTo = Math.Max(newT0, newT1);

            List<Interval1D> cutIntervals = new List<Interval1D>();

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (ex == null)
                    continue;

                if (Math.Abs(ex.Z - newZ) > tol)
                    continue;

                XYZ exDir = CanonicalizeDirection(ex.Dir);
                if (exDir == null)
                    continue;

                if (Math.Abs(Cross2D(dir, exDir)) > tol)
                    continue;

                double exOffset = Dot2D(ex.P0, normal);
                if (Math.Abs(exOffset - newOffset) > tol)
                    continue;

                double exT0 = Dot2D(ex.P0, dir);
                double exT1 = Dot2D(ex.P1, dir);
                double exFrom = Math.Min(exT0, exT1);
                double exTo = Math.Max(exT0, exT1);

                double overlapFrom = Math.Max(newFrom, exFrom);
                double overlapTo = Math.Min(newTo, exTo);

                if (overlapTo - overlapFrom > tol)
                {
                    cutIntervals.Add(new Interval1D
                    {
                        From = overlapFrom,
                        To = overlapTo
                    });
                }
            }

            if (cutIntervals.Count == 0)
                return new List<Line> { newLine };

            List<Interval1D> mergedCuts = MergeIntervals(cutIntervals);
            List<Interval1D> remainingIntervals = SubtractIntervals(
                new Interval1D { From = newFrom, To = newTo },
                mergedCuts);

            List<Line> result = new List<Line>();
            foreach (Interval1D interval in remainingIntervals)
            {
                if (interval.To - interval.From <= tol)
                    continue;

                XYZ rp0 = new XYZ(
                    dir.X * interval.From + normal.X * newOffset,
                    dir.Y * interval.From + normal.Y * newOffset,
                    newZ);

                XYZ rp1 = new XYZ(
                    dir.X * interval.To + normal.X * newOffset,
                    dir.Y * interval.To + normal.Y * newOffset,
                    newZ);

                if (rp0.DistanceTo(rp1) > tol)
                    result.Add(Line.CreateBound(rp0, rp1));
            }

            return result;
        }

        private static List<Interval1D> MergeIntervals(List<Interval1D> intervals)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (intervals == null || intervals.Count == 0)
                return result;

            List<Interval1D> ordered = intervals
                .Where(x => x != null && x.To - x.From > tol)
                .OrderBy(x => x.From)
                .ThenBy(x => x.To)
                .ToList();

            if (ordered.Count == 0)
                return result;

            double curFrom = ordered[0].From;
            double curTo = ordered[0].To;

            for (int i = 1; i < ordered.Count; i++)
            {
                Interval1D next = ordered[i];

                if (next.From <= curTo + tol)
                {
                    curTo = Math.Max(curTo, next.To);
                }
                else
                {
                    result.Add(new Interval1D { From = curFrom, To = curTo });
                    curFrom = next.From;
                    curTo = next.To;
                }
            }

            result.Add(new Interval1D { From = curFrom, To = curTo });
            return result;
        }

        private static List<Interval1D> SubtractIntervals(Interval1D source, List<Interval1D> cuts)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (source == null || source.To - source.From <= tol)
                return result;

            if (cuts == null || cuts.Count == 0)
            {
                result.Add(source);
                return result;
            }

            double cursor = source.From;

            foreach (Interval1D cut in cuts)
            {
                if (cut == null || cut.To - cut.From <= tol)
                    continue;

                if (cut.To <= source.From + tol)
                    continue;

                if (cut.From >= source.To - tol)
                    break;

                double cutFrom = Math.Max(source.From, cut.From);
                double cutTo = Math.Min(source.To, cut.To);

                if (cutFrom > cursor + tol)
                {
                    result.Add(new Interval1D
                    {
                        From = cursor,
                        To = cutFrom
                    });
                }

                cursor = Math.Max(cursor, cutTo);
            }

            if (cursor < source.To - tol)
            {
                result.Add(new Interval1D
                {
                    From = cursor,
                    To = source.To
                });
            }

            return result;
        }

        private static List<Line> MergeCollinearLines(List<Line> lines)
        {
            const double tol = 1e-6;

            List<GenericAxisLineData> data = new List<GenericAxisLineData>();

            foreach (Line line in lines)
            {
                if (line == null || line.Length <= tol)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);

                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                dir = CanonicalizeDirection(dir);

                XYZ normal = new XYZ(-dir.Y, dir.X, 0);
                double offset = Dot2D(p0, normal);
                double z = 0.5 * (p0.Z + p1.Z);

                double t0 = Dot2D(p0, dir);
                double t1 = Dot2D(p1, dir);

                double from = Math.Min(t0, t1);
                double to = Math.Max(t0, t1);

                data.Add(new GenericAxisLineData
                {
                    Dir = dir,
                    Normal = normal,
                    Offset = offset,
                    Z = z,
                    From = from,
                    To = to
                });
            }

            var groups = data
                .GroupBy(x => new GenericAxisGroupKey(x.Dir, x.Offset, x.Z, tol))
                .ToList();

            List<Line> result = new List<Line>();

            foreach (var group in groups)
            {
                List<GenericAxisLineData> ordered = group
                    .OrderBy(x => x.From)
                    .ThenBy(x => x.To)
                    .ToList();

                if (ordered.Count == 0)
                    continue;

                double curFrom = ordered[0].From;
                double curTo = ordered[0].To;

                for (int i = 1; i < ordered.Count; i++)
                {
                    GenericAxisLineData next = ordered[i];
                    if (next.From <= curTo + tol)
                    {
                        curTo = Math.Max(curTo, next.To);
                    }
                    else
                    {
                        result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
                        curFrom = next.From;
                        curTo = next.To;
                    }
                }

                result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
            }

            return result;
        }

        private static Line BuildGenericAxisLine(XYZ dir, double offset, double z, double from, double to)
        {
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            XYZ p0 = new XYZ(
                dir.X * from + normal.X * offset,
                dir.Y * from + normal.Y * offset,
                z);

            XYZ p1 = new XYZ(
                dir.X * to + normal.X * offset,
                dir.Y * to + normal.Y * offset,
                z);

            return Line.CreateBound(p0, p1);
        }

        private static bool TryIntersectInfiniteLines2D(XYZ p1, XYZ d1, XYZ p2, XYZ d2, out XYZ intersection)
        {
            const double tol = 1e-12;

            intersection = null;
            double cross = Cross2D(d1, d2);
            if (Math.Abs(cross) < tol)
                return false;

            XYZ delta = p2 - p1;
            double t = Cross2D(delta, d2) / cross;

            intersection = new XYZ(
                p1.X + d1.X * t,
                p1.Y + d1.Y * t,
                0);

            return true;
        }

        private static bool PointOnSegment2D(XYZ p, XYZ a, XYZ b, double tol)
        {
            double ab = Distance2D(a, b);
            double ap = Distance2D(a, p);
            double pb = Distance2D(p, b);
            return Math.Abs((ap + pb) - ab) <= tol;
        }

        private static bool AreCollinear2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            if (Math.Abs(Cross2D(ad, bd)) > tol)
                return false;

            if (Math.Abs(Cross2D(b0 - a0, ad)) > tol)
                return false;

            return true;
        }

        private static XYZ Normalize2D(XYZ v)
        {
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-12)
                return null;

            return new XYZ(v.X / len, v.Y / len, 0);
        }

        private static XYZ CanonicalizeDirection(XYZ dir)
        {
            if (dir == null)
                return null;

            if (dir.X < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            if (Math.Abs(dir.X) < 1e-9 && dir.Y < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            return new XYZ(dir.X, dir.Y, 0);
        }

        private static double Dot2D(XYZ a, XYZ b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static double GetSignedAreaXY(List<XYZ> pts)
        {
            if (pts == null || pts.Count < 3)
                return 0.0;

            double area2 = 0.0;
            for (int i = 0; i < pts.Count; i++)
            {
                XYZ a = pts[i];
                XYZ b = pts[(i + 1) % pts.Count];
                area2 += (a.X * b.Y) - (b.X * a.Y);
            }

            return area2 * 0.5;
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Revit2024 || Debug2024
            return id.Value;
#else
            return id.IntegerValue;
#endif
        }

        private static double ConvertMmToInternal(int valueMm)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double ConvertInternalToMm(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double RoundTol(double value, double tol)
        {
            return Math.Round(value / tol) * tol;
        }

        private class PreparedApartmentWalls
        {
            public ElementId ApartmentId { get; set; }
            public WallType WallType { get; set; }
            public int ThicknessMm { get; set; }
            public List<Line> AxisLines { get; set; }
        }

        private class OffsetLine2D
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        private class ExistingWallLineInfo
        {
            public ElementId WallId { get; set; }
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        private class Interval1D
        {
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisLineData
        {
            public XYZ Dir { get; set; }
            public XYZ Normal { get; set; }
            public double Offset { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisGroupKey : IEquatable<GenericAxisGroupKey>
        {
            public XYZ Dir { get; private set; }
            public double Offset { get; private set; }
            public double Z { get; private set; }

            public GenericAxisGroupKey(XYZ dir, double offset, double z, double tol)
            {
                Dir = new XYZ(RoundTol(dir.X, tol), RoundTol(dir.Y, tol), 0);
                Offset = RoundTol(offset, tol);
                Z = RoundTol(z, tol);
            }

            public bool Equals(GenericAxisGroupKey other)
            {
                if (other == null)
                    return false;

                return Math.Abs(Dir.X - other.Dir.X) < 1e-9
                    && Math.Abs(Dir.Y - other.Dir.Y) < 1e-9
                    && Math.Abs(Offset - other.Offset) < 1e-9
                    && Math.Abs(Z - other.Z) < 1e-9;
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as GenericAxisGroupKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + Dir.X.GetHashCode();
                    hash = hash * 23 + Dir.Y.GetHashCode();
                    hash = hash * 23 + Offset.GetHashCode();
                    hash = hash * 23 + Z.GetHashCode();
                    return hash;
                }
            }
        }
    }
}