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

                    List<Solid> roomSolids = new List<Solid>();
                    List<string> debugMessages = new List<string>();

                    foreach (FamilyInstance apartmentFi in apartmentInstances)
                    {
                        try
                        {
                            FamilyInstance roomFi = FindRoomSubComponent(doc, apartmentFi);
                            if (roomFi == null)
                            {
                                debugMessages.Add("Не найден вложенный экземпляр 'Помещение' у экземпляра Id=" + GetElementIdValue(apartmentFi.Id));
                                continue;
                            }

                            Solid roomSolid = BuildRoomSolidFromInstance(roomFi, preset.WallHeight);
                            if (roomSolid == null || roomSolid.Volume <= 0)
                            {
                                debugMessages.Add("Не удалось построить solid помещения у экземпляра Id=" + GetElementIdValue(apartmentFi.Id));
                                continue;
                            }

                            roomSolids.Add(roomSolid);
                        }
                        catch (Exception exRoom)
                        {
                            debugMessages.Add("Ошибка обработки экземпляра Id=" + GetElementIdValue(apartmentFi.Id) + ": " + exRoom.Message);
                        }
                    }

                    if (roomSolids.Count == 0)
                    {
                        string debugText = debugMessages.Count > 0
                            ? "\n\n" + string.Join("\n", debugMessages)
                            : "";
                        TaskDialog.Show("Apartment Manager", "Не удалось получить ни одного solid для построения стен." + debugText);
                        return Result.Cancelled;
                    }

                    Solid unitedSolid;
                    try
                    {
                        unitedSolid = UnionSolids(roomSolids);
                    }
                    catch (Exception exUnion)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось объединить контуры помещений:\n" + exUnion.Message);
                        return Result.Failed;
                    }

                    List<CurveLoop> outerLoops;
                    try
                    {
                        outerLoops = GetOuterBottomFaceLoops(unitedSolid);
                    }
                    catch (Exception exLoops)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось получить внешний контур объединённого помещения:\n" + exLoops.Message);
                        return Result.Failed;
                    }

                    if (outerLoops == null || outerLoops.Count == 0)
                    {
                        TaskDialog.Show("Ошибка", "Не найден внешний контур для построения стен.");
                        return Result.Failed;
                    }

                    double wallHeightInternal = ConvertMmToInternal(preset.WallHeight);

                    using (Transaction t = new Transaction(doc, "KPLN - Построение стен по внешнему контуру"))
                    {
                        t.Start();

                        CreateWallsByLoops(
                            doc,
                            outerLoops,
                            floorPlan.GenLevel,
                            wallType,
                            wallHeightInternal);

                        t.Commit();
                    }

                    string report =
                        "Построение завершено.\n" +
                        "Экземпляров квартир: " + apartmentInstances.Count + "\n" +
                        "Solid помещений: " + roomSolids.Count + "\n" +
                        "Внешних контуров: " + outerLoops.Count;

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

        private static FamilyInstance FindRoomSubComponent(Document doc, FamilyInstance apartmentInstance)
        {
            if (apartmentInstance == null)
                return null;

            ICollection<ElementId> subIds = apartmentInstance.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return null;

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
                    return subFi;
            }

            return null;
        }

        private static Solid BuildRoomSolidFromInstance(FamilyInstance roomFi, int wallHeightMm)
        {
            if (roomFi == null)
                throw new ArgumentNullException("roomFi");

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");
            double height = ConvertMmToInternal(wallHeightMm);

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                throw new Exception("Не удалось получить Transform для вложенного помещения.");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            // Здесь заложено предположение, что точка вставки вложенного "Помещения" находится в центре.
            // Если у семейства база не по центру, нужно будет заменить локальные точки.
            XYZ p1 = new XYZ(-halfW, -halfD, 0);
            XYZ p2 = new XYZ(halfW, -halfD, 0);
            XYZ p3 = new XYZ(halfW, halfD, 0);
            XYZ p4 = new XYZ(-halfW, halfD, 0);

            List<Curve> profile = new List<Curve>();
            profile.Add(Line.CreateBound(p1, p2));
            profile.Add(Line.CreateBound(p2, p3));
            profile.Add(Line.CreateBound(p3, p4));
            profile.Add(Line.CreateBound(p4, p1));

            CurveLoop loop = CurveLoop.Create(profile);
            List<CurveLoop> loops = new List<CurveLoop> { loop };

            Solid localSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                loops,
                XYZ.BasisZ,
                height);

            return SolidUtils.CreateTransformed(localSolid, tr);
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

        private static Solid UnionSolids(List<Solid> solids)
        {
            if (solids == null || solids.Count == 0)
                throw new Exception("Список solids пуст.");

            Solid result = solids[0];

            for (int i = 1; i < solids.Count; i++)
            {
                result = BooleanOperationsUtils.ExecuteBooleanOperation(
                    result,
                    solids[i],
                    BooleanOperationsType.Union);
            }

            return result;
        }

        private static List<CurveLoop> GetOuterBottomFaceLoops(Solid solid)
        {
            if (solid == null)
                throw new ArgumentNullException("solid");

            double minZ = double.MaxValue;
            List<PlanarFace> bottomFaces = new List<PlanarFace>();

            foreach (Face face in solid.Faces)
            {
                PlanarFace pf = face as PlanarFace;
                if (pf == null)
                    continue;

                XYZ normal = pf.FaceNormal;
                if (normal == null)
                    continue;

                if (!IsAlmostEqual(normal, -XYZ.BasisZ))
                    continue;

                double z = pf.Origin.Z;

                if (z < minZ - 1e-9)
                {
                    minZ = z;
                    bottomFaces.Clear();
                    bottomFaces.Add(pf);
                }
                else if (Math.Abs(z - minZ) < 1e-9)
                {
                    bottomFaces.Add(pf);
                }
            }

            if (bottomFaces.Count == 0)
                throw new Exception("Не найдена нижняя горизонтальная грань у объединённого solid.");

            List<CurveLoop> result = new List<CurveLoop>();

            foreach (PlanarFace bottomFace in bottomFaces)
            {
                CurveLoop outerLoop = GetLongestLoop(bottomFace);
                if (outerLoop != null)
                    result.Add(outerLoop);
            }

            return result;
        }

        private static CurveLoop GetLongestLoop(PlanarFace face)
        {
            if (face == null)
                return null;

            EdgeArrayArray edgeLoops = face.EdgeLoops;
            CurveLoop bestLoop = null;
            double bestLength = -1.0;

            foreach (EdgeArray edgeArray in edgeLoops)
            {
                List<Curve> curves = new List<Curve>();
                double totalLength = 0.0;

                foreach (Edge edge in edgeArray)
                {
                    Curve c = edge.AsCurve();
                    curves.Add(c);
                    totalLength += c.Length;
                }

                CurveLoop loop = CurveLoop.Create(curves);
                if (totalLength > bestLength)
                {
                    bestLength = totalLength;
                    bestLoop = loop;
                }
            }

            return bestLoop;
        }

        private static void CreateWallsByLoops(
            Document doc,
            IEnumerable<CurveLoop> loops,
            Level level,
            WallType wallType,
            double wallHeight)
        {
            foreach (CurveLoop loop in loops)
            {
                foreach (Curve curve in loop)
                {
                    Line line = curve as Line;
                    if (line == null)
                        continue;

                    Wall.Create(
                        doc,
                        line,
                        wallType.Id,
                        level.Id,
                        wallHeight,
                        0,
                        false,
                        false);
                }
            }
        }

        private static bool IsAlmostEqual(XYZ a, XYZ b, double tol = 1e-9)
        {
            if (a == null || b == null)
                return false;

            return a.DistanceTo(b) < tol;
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
    }
}