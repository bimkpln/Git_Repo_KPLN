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

                    WallType wallType = FindWallTypeByName(doc, preset.WallType);
                    if (wallType == null)
                    {
                        TaskDialog.Show("Ошибка", "Не найден тип стены: " + preset.WallType);
                        return Result.Failed;
                    }

                    List<FamilyInstance> apartmentInstances = GetPlacedApartmentInstancesOnLevel(doc, floorPlan.GenLevel);
                    if (apartmentInstances.Count == 0)
                    {
                        TaskDialog.Show("Apartment Manager", "На текущем уровне не найдено ранее размещённых экземпляров квартир.");
                        return Result.Cancelled;
                    }

                    List<AxisRoomEdge> allEdges = new List<AxisRoomEdge>();
                    List<string> debugMessages = new List<string>();

                    foreach (FamilyInstance apartmentFi in apartmentInstances)
                    {
                        List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
                        if (roomInstances.Count == 0)
                        {
                            debugMessages.Add("Не найдены вложенные экземпляры 'Помещение' у экземпляра Id=" + GetElementIdValue(apartmentFi.Id));
                            continue;
                        }

                        foreach (FamilyInstance roomFi in roomInstances)
                        {
                            try
                            {
                                CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi);
                                List<AxisRoomEdge> roomEdges = ExtractAxisEdgesFromLoop(roomLoop, GetElementIdValue(roomFi.Id));
                                allEdges.AddRange(roomEdges);
                            }
                            catch (Exception exRoom)
                            {
                                debugMessages.Add("Ошибка обработки вложенного помещения Id=" + GetElementIdValue(roomFi.Id) + ": " + exRoom.Message);
                            }
                        }
                    }

                    if (allEdges.Count == 0)
                    {
                        string debugText = debugMessages.Count > 0
                            ? "\n\n" + string.Join("\n", debugMessages)
                            : "";
                        TaskDialog.Show("Apartment Manager", "Не удалось получить ни одного контура помещений." + debugText);
                        return Result.Cancelled;
                    }

                    List<Line> wallAxisLines = BuildWallAxisLinesFromRoomEdges(allEdges, wallType.Width);
                    if (wallAxisLines.Count == 0)
                    {
                        string debugText = debugMessages.Count > 0
                            ? "\n\n" + string.Join("\n", debugMessages)
                            : "";
                        TaskDialog.Show("Apartment Manager", "Не удалось вычислить оси стен." + debugText);
                        return Result.Cancelled;
                    }

                    double wallHeightInternal = ConvertMmToInternal(preset.WallHeight);

                    using (Transaction t = new Transaction(doc, "KPLN - Построение стен по помещениям"))
                    {
                        t.Start();

                        foreach (Line axis in wallAxisLines)
                        {
                            if (axis == null || axis.Length < 1e-6)
                                continue;

                            Wall.Create(
                                doc,
                                axis,
                                wallType.Id,
                                floorPlan.GenLevel.Id,
                                wallHeightInternal,
                                0,
                                false,
                                false);
                        }

                        t.Commit();
                    }

                    string report =
                        "Построение завершено.\n" +
                        "Экземпляров квартир: " + apartmentInstances.Count + "\n" +
                        "Границ помещений: " + allEdges.Count + "\n" +
                        "Стеновых осей: " + wallAxisLines.Count;

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

            // Если база семейства "Помещение" не по центру, а из угла,
            // замени точки на:
            // p1=(0,0,0), p2=(width,0,0), p3=(width,depth,0), p4=(0,depth,0)
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

        private static List<AxisRoomEdge> ExtractAxisEdgesFromLoop(CurveLoop loop, long ownerId)
        {
            if (loop == null)
                throw new ArgumentNullException("loop");

            List<Line> lines = loop.OfType<Curve>().Select(x => x as Line).Where(x => x != null).ToList();
            if (lines.Count == 0)
                throw new Exception("Контур помещения не содержит линейных сегментов.");

            double area = GetSignedAreaXY(loop);
            bool ccw = area > 0.0;

            List<AxisRoomEdge> result = new List<AxisRoomEdge>();

            foreach (Line line in lines)
            {
                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);

                if (Math.Abs(p0.Y - p1.Y) < 1e-6)
                {
                    double y = 0.5 * (p0.Y + p1.Y);
                    double z = 0.5 * (p0.Z + p1.Z);
                    double from = Math.Min(p0.X, p1.X);
                    double to = Math.Max(p0.X, p1.X);

                    double dx = p1.X - p0.X;
                    int outwardSign = GetSign(-dx);
                    if (!ccw)
                        outwardSign = -outwardSign;

                    result.Add(new AxisRoomEdge
                    {
                        Orientation = AxisOrientation.Horizontal,
                        Coord = y,
                        Z = z,
                        From = from,
                        To = to,
                        OutwardSign = outwardSign,
                        OwnerId = ownerId
                    });
                }
                else if (Math.Abs(p0.X - p1.X) < 1e-6)
                {
                    double x = 0.5 * (p0.X + p1.X);
                    double z = 0.5 * (p0.Z + p1.Z);
                    double from = Math.Min(p0.Y, p1.Y);
                    double to = Math.Max(p0.Y, p1.Y);

                    double dy = p1.Y - p0.Y;
                    int outwardSign = GetSign(dy);
                    if (!ccw)
                        outwardSign = -outwardSign;

                    result.Add(new AxisRoomEdge
                    {
                        Orientation = AxisOrientation.Vertical,
                        Coord = x,
                        Z = z,
                        From = from,
                        To = to,
                        OutwardSign = outwardSign,
                        OwnerId = ownerId
                    });
                }
                else
                {
                    throw new Exception("Обнаружен неортогональный сегмент. Этот код рассчитан на прямоугольные помещения.");
                }
            }

            return result;
        }

        private static List<Line> BuildWallAxisLinesFromRoomEdges(List<AxisRoomEdge> sourceEdges, double wallWidth)
        {
            const double tol = 1e-6;
            double halfWidth = wallWidth / 2.0;
            double maxPairDistance = wallWidth + tol;

            List<AxisRoomEdge> working = sourceEdges
                .Where(x => x != null && x.To - x.From > tol)
                .Select(x => x.Clone())
                .ToList();

            List<Line> axisLines = new List<Line>();

            bool foundPair;
            do
            {
                foundPair = false;

                int bestI = -1;
                int bestJ = -1;
                double bestDistance = double.MaxValue;
                double bestFrom = 0.0;
                double bestTo = 0.0;

                for (int i = 0; i < working.Count; i++)
                {
                    AxisRoomEdge a = working[i];
                    if (a == null || a.Length <= tol)
                        continue;

                    for (int j = i + 1; j < working.Count; j++)
                    {
                        AxisRoomEdge b = working[j];
                        if (b == null || b.Length <= tol)
                            continue;

                        if (a.Orientation != b.Orientation)
                            continue;

                        if (Math.Abs(a.Z - b.Z) > tol)
                            continue;

                        if (a.OutwardSign == b.OutwardSign)
                            continue;

                        double overlapFrom = Math.Max(a.From, b.From);
                        double overlapTo = Math.Min(a.To, b.To);

                        if (overlapTo - overlapFrom <= tol)
                            continue;

                        double distance = Math.Abs(a.Coord - b.Coord);
                        if (distance > maxPairDistance)
                            continue;

                        if (distance < bestDistance)
                        {
                            bestDistance = distance;
                            bestI = i;
                            bestJ = j;
                            bestFrom = overlapFrom;
                            bestTo = overlapTo;
                            foundPair = true;
                        }
                    }
                }

                if (foundPair)
                {
                    AxisRoomEdge a = working[bestI];
                    AxisRoomEdge b = working[bestJ];

                    double axisCoord = 0.5 * (a.Coord + b.Coord);
                    axisLines.Add(BuildAxisLine(a.Orientation, axisCoord, a.Z, bestFrom, bestTo));

                    List<AxisRoomEdge> aParts = SubtractInterval(a, bestFrom, bestTo);
                    List<AxisRoomEdge> bParts = SubtractInterval(b, bestFrom, bestTo);

                    if (bestI > bestJ)
                    {
                        working.RemoveAt(bestI);
                        working.RemoveAt(bestJ);
                    }
                    else
                    {
                        working.RemoveAt(bestJ);
                        working.RemoveAt(bestI);
                    }

                    working.AddRange(aParts);
                    working.AddRange(bParts);
                }
            }
            while (foundPair);

            foreach (AxisRoomEdge edge in working)
            {
                if (edge == null || edge.Length <= tol)
                    continue;

                double axisCoord = edge.Coord + edge.OutwardSign * halfWidth;
                axisLines.Add(BuildAxisLine(edge.Orientation, axisCoord, edge.Z, edge.From, edge.To));
            }

            return MergeAxisAlignedLines(axisLines);
        }

        private static List<AxisRoomEdge> SubtractInterval(AxisRoomEdge edge, double cutFrom, double cutTo)
        {
            const double tol = 1e-6;
            List<AxisRoomEdge> result = new List<AxisRoomEdge>();

            if (edge == null)
                return result;

            if (cutFrom > edge.From + tol)
            {
                result.Add(new AxisRoomEdge
                {
                    Orientation = edge.Orientation,
                    Coord = edge.Coord,
                    Z = edge.Z,
                    From = edge.From,
                    To = cutFrom,
                    OutwardSign = edge.OutwardSign,
                    OwnerId = edge.OwnerId
                });
            }

            if (cutTo < edge.To - tol)
            {
                result.Add(new AxisRoomEdge
                {
                    Orientation = edge.Orientation,
                    Coord = edge.Coord,
                    Z = edge.Z,
                    From = cutTo,
                    To = edge.To,
                    OutwardSign = edge.OutwardSign,
                    OwnerId = edge.OwnerId
                });
            }

            return result;
        }

        private static List<Line> MergeAxisAlignedLines(List<Line> lines)
        {
            const double tol = 1e-6;

            List<AxisLineData> horizontal = new List<AxisLineData>();
            List<AxisLineData> vertical = new List<AxisLineData>();
            List<Line> other = new List<Line>();

            foreach (Line line in lines)
            {
                if (line == null)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);

                if (Math.Abs(p0.Y - p1.Y) < tol)
                {
                    double y = RoundTol(p0.Y, tol);
                    double z = RoundTol((p0.Z + p1.Z) * 0.5, tol);
                    double from = Math.Min(p0.X, p1.X);
                    double to = Math.Max(p0.X, p1.X);

                    horizontal.Add(new AxisLineData
                    {
                        Orientation = AxisOrientation.Horizontal,
                        Coord = y,
                        Z = z,
                        From = from,
                        To = to
                    });
                }
                else if (Math.Abs(p0.X - p1.X) < tol)
                {
                    double x = RoundTol(p0.X, tol);
                    double z = RoundTol((p0.Z + p1.Z) * 0.5, tol);
                    double from = Math.Min(p0.Y, p1.Y);
                    double to = Math.Max(p0.Y, p1.Y);

                    vertical.Add(new AxisLineData
                    {
                        Orientation = AxisOrientation.Vertical,
                        Coord = x,
                        Z = z,
                        From = from,
                        To = to
                    });
                }
                else
                {
                    other.Add(line);
                }
            }

            List<Line> result = new List<Line>();
            result.AddRange(MergeAxisGroup(horizontal));
            result.AddRange(MergeAxisGroup(vertical));
            result.AddRange(other);

            return result;
        }

        private static List<Line> MergeAxisGroup(List<AxisLineData> items)
        {
            const double tol = 1e-6;
            List<Line> result = new List<Line>();

            var grouped = items
                .GroupBy(x => new AxisGroupKey(x.Orientation, x.Coord, x.Z, tol))
                .ToList();

            foreach (var group in grouped)
            {
                List<AxisLineData> ordered = group
                    .OrderBy(x => x.From)
                    .ThenBy(x => x.To)
                    .ToList();

                if (ordered.Count == 0)
                    continue;

                double curFrom = ordered[0].From;
                double curTo = ordered[0].To;

                for (int i = 1; i < ordered.Count; i++)
                {
                    AxisLineData next = ordered[i];

                    if (next.From <= curTo + tol)
                    {
                        curTo = Math.Max(curTo, next.To);
                    }
                    else
                    {
                        result.Add(BuildAxisLine(group.Key.Orientation, group.Key.Coord, group.Key.Z, curFrom, curTo));
                        curFrom = next.From;
                        curTo = next.To;
                    }
                }

                result.Add(BuildAxisLine(group.Key.Orientation, group.Key.Coord, group.Key.Z, curFrom, curTo));
            }

            return result;
        }

        private static Line BuildAxisLine(AxisOrientation orientation, double coord, double z, double from, double to)
        {
            if (orientation == AxisOrientation.Horizontal)
            {
                XYZ p0 = new XYZ(from, coord, z);
                XYZ p1 = new XYZ(to, coord, z);
                return Line.CreateBound(p0, p1);
            }
            else
            {
                XYZ p0 = new XYZ(coord, from, z);
                XYZ p1 = new XYZ(coord, to, z);
                return Line.CreateBound(p0, p1);
            }
        }

        private static double GetSignedAreaXY(CurveLoop loop)
        {
            if (loop == null)
                return 0.0;

            List<XYZ> pts = new List<XYZ>();

            foreach (Curve c in loop)
            {
                if (c == null)
                    continue;

                XYZ p = c.GetEndPoint(0);
                if (pts.Count == 0 || !IsAlmostEqual(pts[pts.Count - 1], p))
                    pts.Add(p);
            }

            if (pts.Count < 3)
                return 0.0;

            if (!IsAlmostEqual(pts[0], pts[pts.Count - 1]))
                pts.Add(pts[0]);

            double area2 = 0.0;
            for (int i = 0; i < pts.Count - 1; i++)
            {
                XYZ a = pts[i];
                XYZ b = pts[i + 1];
                area2 += (a.X * b.Y) - (b.X * a.Y);
            }

            return area2 * 0.5;
        }

        private static int GetSign(double value)
        {
            if (value > 1e-9) return 1;
            if (value < -1e-9) return -1;
            return 0;
        }

        private static bool IsAlmostEqual(XYZ a, XYZ b, double tol = 1e-9)
        {
            if (a == null || b == null)
                return false;

            return a.DistanceTo(b) < tol;
        }

        private static double RoundTol(double value, double tol)
        {
            return Math.Round(value / tol) * tol;
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

        private enum AxisOrientation
        {
            Horizontal,
            Vertical
        }

        private class AxisRoomEdge
        {
            public AxisOrientation Orientation { get; set; }
            public double Coord { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
            public int OutwardSign { get; set; }
            public long OwnerId { get; set; }

            public double Length
            {
                get { return To - From; }
            }

            public AxisRoomEdge Clone()
            {
                return new AxisRoomEdge
                {
                    Orientation = Orientation,
                    Coord = Coord,
                    Z = Z,
                    From = From,
                    To = To,
                    OutwardSign = OutwardSign,
                    OwnerId = OwnerId
                };
            }
        }

        private class AxisLineData
        {
            public AxisOrientation Orientation { get; set; }
            public double Coord { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
        }

        private class AxisGroupKey : IEquatable<AxisGroupKey>
        {
            public AxisOrientation Orientation { get; private set; }
            public double Coord { get; private set; }
            public double Z { get; private set; }

            public AxisGroupKey(AxisOrientation orientation, double coord, double z, double tol)
            {
                Orientation = orientation;
                Coord = RoundTol(coord, tol);
                Z = RoundTol(z, tol);
            }

            public bool Equals(AxisGroupKey other)
            {
                if (other == null)
                    return false;

                return Orientation == other.Orientation
                    && Math.Abs(Coord - other.Coord) < 1e-9
                    && Math.Abs(Z - other.Z) < 1e-9;
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as AxisGroupKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + Orientation.GetHashCode();
                    hash = hash * 23 + Coord.GetHashCode();
                    hash = hash * 23 + Z.GetHashCode();
                    return hash;
                }
            }
        }
    }
}