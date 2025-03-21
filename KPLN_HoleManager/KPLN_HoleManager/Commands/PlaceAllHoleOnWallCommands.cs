using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_HoleManager.Forms;
using System.Windows.Forms;

namespace KPLN_HoleManager.Commands
{
    public class PlaceAllHoleOnWallCommand
    {
        private readonly string _userFullName;
        private readonly string _departmentName;
        private readonly Element _selectedWall;

        private string _departmentHoleName;
        private string _sendingDepartmentHoleName;
        private string _holeTypeName;

        private bool transactionStatus;
        private bool deleteLastHole;

        public PlaceAllHoleOnWallCommand(string userFullName, string departmentName, Element selectedElement, string departmentHoleName, string sendingDepartmentHoleName, string holeTypeName)
        {
            _userFullName = userFullName;
            _departmentName = departmentName;
            _selectedWall = selectedElement;
            _departmentHoleName = departmentHoleName;
            _sendingDepartmentHoleName = sendingDepartmentHoleName;
            _holeTypeName = holeTypeName;
        }

        public static void Execute(UIApplication uiApp, string userFullName, string departmentName, Element selectedElement, string departmentHoleName, string sendingDepartmentHoleName, string holeTypeName)
        {
            if (uiApp == null)
            {
                TaskDialog.Show("Ошибка", "Revit API не доступен.");
                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                return;
            }

            // Создаём экземпляр команды и передаём её в ExternalEvent
            var command = new PlaceAllHoleOnWallCommand(userFullName, departmentName, selectedElement, departmentHoleName, sendingDepartmentHoleName, holeTypeName);
            _ExternalEventHandler.Instance.Raise((app) => command.Run(app));
        }

        public void Run(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Чтение файлов преднастроек
            List<string> settings = DockableManagerFormSettings.LoadSettings();
            if (settings != null)
            {
                _departmentHoleName = settings[2];
                _sendingDepartmentHoleName = settings[3];
                _holeTypeName = settings[4];
            }

            try
            {
                /// Стена
                Wall wall = _selectedWall as Wall;
                LocationCurve locationCurve = wall.Location as LocationCurve;
               
                /// Элементы пересекающие стену. Все файлы
                List<Element> mainFileElements = new List<Element>();
                List<Element> linkedFileElements = new List<Element>();

                BuiltInCategory[] categories = new BuiltInCategory[]
                {
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit
                };

                ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(categories);

                // Основной файл
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WherePasses(categoryFilter).WhereElementIsNotElementType();
                mainFileElements.AddRange(collector.ToList());

                // Связные файлы            
                FilteredElementCollector linkedCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance));

                Dictionary<Element, Autodesk.Revit.DB.Transform> linkedElementTransforms = new Dictionary<Element, Autodesk.Revit.DB.Transform>();

                foreach (RevitLinkInstance linkInstance in linkedCollector)
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null) continue;

                    Autodesk.Revit.DB.Transform linkTransform = linkInstance.GetTotalTransform();

                    FilteredElementCollector linkedElements = new FilteredElementCollector(linkedDoc)
                        .WherePasses(categoryFilter)
                        .WhereElementIsNotElementType();

                    foreach (var el in linkedElements)
                    {
                        linkedFileElements.Add(el);
                        linkedElementTransforms[el] = linkTransform;
                    }
                }

                /// Элементы пересекающие стену. Поиск нужных элементов с учётом фильтроа
                List<Element> intersectingElements = new List<Element>();

                if (IsWallFromLinkedDocument(doc, wall))
                {
                    if (_departmentHoleName == "АР")
                    {
                        TaskDialog.Show("Ошибка", "Вы пытаетесь вырезать отверстие в стене из связанного файла. Операция отменена.");
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        return;
                    }

                    Autodesk.Revit.DB.Transform wallTransform = FindLinkTransformForElement(doc, wall);
                    if (wallTransform == null)
                        throw new Exception("Не удалось найти трансформацию связного файла для стены.");

                    foreach (var el in mainFileElements)
                    {
                        if (ElementsIntersect(wall, el, null, wallTransform))
                            intersectingElements.Add(el);
                    }
                }
                else
                {
                    foreach (var el in linkedFileElements)
                    {
                        Autodesk.Revit.DB.Transform transform = linkedElementTransforms.ContainsKey(el) ? linkedElementTransforms[el] : null;

                        if (ElementsIntersect(wall, el, transform))
                            intersectingElements.Add(el);
                    }
                }

                // Определение файла семейства
                string familyFileName = GetFamilyFileName(_departmentHoleName, _holeTypeName);

                if (string.IsNullOrEmpty(familyFileName))
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить файл семейства.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                string pluginFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string familyPath = Path.Combine(pluginFolder, "Families", familyFileName);

                if (!File.Exists(familyPath))
                {
                    TaskDialog.Show("Ошибка", $"Файл семейства не найден:\n{familyPath}");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                // Запуск общей транзакции
                using (Transaction tx = new Transaction(doc, $"KPLN. Разместить отверстия на стене {wall.Id}"))
                {
                    tx.Start();
                    transactionStatus = true;

                    // Загрузка семейства
                    Family family = LoadOrGetExistingFamily(doc, familyPath);

                    if (family == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        return;
                    }

                    FamilySymbol holeSymbol = GetFamilySymbol(family, doc);
                    if (holeSymbol == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Не удалось получить типоразмер семейства.");
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        return;
                    }

                    if (!holeSymbol.IsActive)
                    {
                        holeSymbol.Activate();
                        doc.Regenerate();
                    }

                    foreach (Element cElement in intersectingElements)
                    {
                        List<XYZ> holeLocation = GetIntersectionPoint(doc, _selectedWall, cElement, _departmentHoleName, _holeTypeName);

                        foreach (XYZ cHoleLocation in holeLocation)
                        {
                            // Создаём отверстие без преднастроек (указывается только точка самого отверстия)
                            FamilyInstance holeInstance = doc.Create.NewFamilyInstance(cHoleLocation, holeSymbol, _selectedWall, StructuralType.NonStructural);

                            if (holeInstance == null)
                            {
                                tx.RollBack();
                                TaskDialog.Show("Ошибка", "Ошибка при создании отверстия.");
                                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                                return;
                            }







                        }






                        string message = $"Найдено точек: {holeLocation.Count}\n\n";
                        foreach (XYZ point in holeLocation)
                        {
                            message += $"X: {point.X:F2}, Y: {point.Y:F2}, Z: {point.Z:F2}\n";
                        }
                        TaskDialog.Show("Координаты отверстий", message);
                    }
               
                    tx.Commit();
                }







                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
            }          
        }









        /// <summary>
        /// Стена. Определение типа стены
        /// </summary>
        public bool IsWallFromLinkedDocument(Document mainDocument, Wall wall)
        {
            return wall != null && !wall.Document.Equals(mainDocument);
        }

        /// <summary>
        /// Вспомогательный метод. Метод применения Transform для элементов из связных файлов
        /// </summary>
        private Autodesk.Revit.DB.Transform FindLinkTransformForElement(Document mainDoc, Element element)
        {
            FilteredElementCollector linkCollector = new FilteredElementCollector(mainDoc).OfClass(typeof(RevitLinkInstance));
            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc != null && linkedDoc.Equals(element.Document))
                    return linkInstance.GetTotalTransform();
            }
            return null;
        }

        /// <summary>
        /// Вспомогательный метод. Стена. Поиск пересечения элементов и стены
        /// </summary>
        private bool ElementsIntersect(Element wall, Element otherElement, 
            Autodesk.Revit.DB.Transform otherTransform = null, Autodesk.Revit.DB.Transform wallTransform = null)
        {
            Solid wallSolid = GetElementSolid(wall, wallTransform);
            Solid otherSolid = GetElementSolid(otherElement, otherTransform);

            if (wallSolid == null || otherSolid == null)
                return false;

            try
            {
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                    wallSolid, otherSolid, BooleanOperationsType.Intersect);

                return intersection != null && intersection.Volume > 1e-6;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Вспомогательный метод. Обработка геометрии
        /// </summary>
        private Solid GetElementSolid(Element element, Autodesk.Revit.DB.Transform transform = null)
        {
            Options options = new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geoElement = element.get_Geometry(options);
            if (geoElement == null) return null;

            foreach (GeometryObject obj in geoElement)
            {
                Solid solid = null;

                if (obj is GeometryInstance geoInst)
                {
                    foreach (GeometryObject instObj in geoInst.GetSymbolGeometry())
                    {
                        if (instObj is Solid s && s.Volume > 0)
                        {
                            solid = s;
                            break;
                        }
                    }
                }
                else if (obj is Solid s && s.Volume > 0)
                {
                    solid = s;
                }

                if (solid != null)
                {
                    if (transform != null)
                        solid = SolidUtils.CreateTransformed(solid, transform);

                    return solid;
                }
            }

            return null;
        }

        /// <summary>
        /// Семейство. Метод выбора файла семейства
        /// </summary>
        private string GetFamilyFileName(string department, string holeType)
        {
            if (department == "ОВиК" || department == "ВК"
                || department == "ЭОМ" || department == "СС")
            {
                department = "ИОС";
            }

            var familyMap = new Dictionary<string, string>
            {
                { "АР_SquareHole", "199_Отверстие прямоугольное_(Об_Стена).rfa" },
                { "АР_RoundHole", "199_Отверстие круглое_(Об_Стена).rfa" },
                { "КР_SquareHole", null },
                { "КР_RoundHole", null },
                { "ИОС_SquareHole", "501_ЗИ_Отверстие_Прямоугольное_Стена_(Об).rfa" },
                { "ИОС_RoundHole", "501_ЗИ_Отверстие_Круглое_Стена_(Об).rfa" }
            };

            string key = $"{department}_{holeType}";
            return familyMap.ContainsKey(key) ? familyMap[key] : null;
        }

        /// <summary>
        /// Семейство. Метод загрузки семейства
        /// </summary>
        private Family LoadOrGetExistingFamily(Document doc, string familyPath)
        {
            string familyName = Path.GetFileNameWithoutExtension(familyPath);

            Family family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name == familyName);

            if (family == null)
            {
                doc.LoadFamily(familyPath, new ReplaceNestedFamilyHandler(), out family);
            }

            return family;
        }

        /// <summary>
        /// Семейство. Метод загрузки типоразмера семейства
        /// </summary>
        private FamilySymbol GetFamilySymbol(Family family, Document doc)
        {
            Dictionary<string, string> departmentMapping = new Dictionary<string, string>
            {
                { "ОВиК", "ОВ" },
                { "ВК", "ВК" },
                { "ЭОМ", "ЭОМ" },
                { "СС", "СС" }
            };

            string targetType = departmentMapping.ContainsKey(_departmentHoleName)
                ? departmentMapping[_departmentHoleName]
                : _departmentHoleName;

            foreach (ElementId id in family.GetFamilySymbolIds())
            {
                FamilySymbol symbol = doc.GetElement(id) as FamilySymbol;
                if (symbol != null && symbol.Name == targetType)
                {
                    return symbol;
                }
            }

            // Если не найден нужный тип, возвращаем первый попавшийся
            foreach (ElementId id in family.GetFamilySymbolIds())
            {
                return doc.GetElement(id) as FamilySymbol;
            }

            return null;
        }


        /// <summary>
        /// XYZ. Метод нахождения точкек пересечения стены и входящего элемента.
        /// </summary>
        private List<XYZ> GetIntersectionPoint(Document doc, Element wall, Element cElement, string department, string holeType)
        {
            List<XYZ> allIntersections = new List<XYZ>();

            // Трансформация для обоих элементов
            Autodesk.Revit.DB.Transform wallTransform = GetElementTransform(doc, wall);
            Autodesk.Revit.DB.Transform cElementTransform = GetElementTransform(doc, cElement);

            // Работа со стеной
            RevitLinkInstance linkInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .FirstOrDefault(link => link.GetLinkDocument()?.Equals(wall.Document) == true);

            Autodesk.Revit.DB.Transform linkTransform = linkInstance?.GetTotalTransform() ?? Autodesk.Revit.DB.Transform.Identity;
            double wallThickness = GetWallThickness(wall);

            // Работа с пересекающим стену элементом
            LocationCurve cElementLocation = cElement.Location as LocationCurve;
            if (cElementLocation == null)
                return null;

            Curve cElementCurve = cElementLocation.Curve;
            List<XYZ> segmentPoints = cElementCurve.Tessellate() as List<XYZ>;
            if (segmentPoints == null || segmentPoints.Count < 2)
                return null;

            // Применяем трансформацию к точкам сегментации, если элемент из связанного файла
            segmentPoints = segmentPoints.Select(pt => cElementTransform.OfPoint(pt)).ToList();

            // 3D-вид и ReferenceIntersector
            bool createdView = false;
            View3D view3D = null;

            // Стена не в связном файле
            if (linkInstance == null)
            {
                view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);

                if (view3D == null)
                {
                    ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.ThreeDimensional);

                    if (viewFamilyType != null)
                    {
                        view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
                        createdView = true;
                    }
                }

                if (view3D == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось создать 3D-вид для ReferenceIntersector.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return null;
                }

                ReferenceIntersector intersector = new ReferenceIntersector(wall.Id, FindReferenceTarget.Face, view3D);

                // Проходим по сегментам кривой
                for (int i = 0; i < segmentPoints.Count - 1; i++)
                {
                    XYZ start = segmentPoints[i];
                    XYZ end = segmentPoints[i + 1];

                    List<XYZ> segmentIntersections = GetIntersectionsForSegment(intersector, start, end, wall, wallThickness);
                    allIntersections.AddRange(segmentIntersections);
                }
            }
            // Стена в связнеом файле
            else
            {
                GeometryElement wallGeometry = wall.get_Geometry(new Options());

                if (wallGeometry == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось получить геометрию стены.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return null;
                }

                RevitLinkInstance linkInstanceWall = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .First(link => link.GetLinkDocument().Equals(wall.Document));

                Autodesk.Revit.DB.Transform linkTransformWall = linkInstance.GetTotalTransform();
                List<Solid> transformedSolids = new List<Solid>();

                foreach (GeometryObject geomObj in wallGeometry)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        transformedSolids.Add(SolidUtils.CreateTransformed(solid, linkTransform));
                    }
                }

                if (transformedSolids.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "Не удалось трансформировать геометрию стены.");
                    return null;
                }

                for (int i = 0; i < segmentPoints.Count - 1; i++)
                {
                    XYZ start = segmentPoints[i];
                    XYZ end = segmentPoints[i + 1];
                    Line segmentLine = Line.CreateBound(start, end);

                    List<XYZ> segmentIntersections = new List<XYZ>();

                    foreach (Solid solid in transformedSolids)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            IntersectionResultArray results;
                            if (face.Intersect(segmentLine, out results) == SetComparisonResult.Overlap && results.Size > 0)
                            {
                                for (int j = 0; j < results.Size; j++)
                                {
                                    XYZ intersectionPoint = results.get_Item(j).XYZPoint;
                                    segmentIntersections.Add(intersectionPoint);
                                }
                            }
                        }
                    }
                    if (segmentIntersections.Count > 0)
                    {
                        // Временное решение расчёта гибких элементов (1 из 2)
                        if (cElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexPipeCurves ||
                            cElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexDuctCurves)
                        {
                            allIntersections.AddRange(segmentIntersections);
                        }

                        else if (segmentIntersections.Count == 1)
                        {
                            allIntersections.Add(segmentIntersections[0]);
                        }
                        else if (segmentIntersections.Count == 2)
                        {
                            XYZ first = segmentIntersections.First();
                            XYZ last = segmentIntersections.Last();
                            XYZ middle = (first + last) / 2;
                            allIntersections.Add(middle);
                        }
                        else
                        {
                            List<XYZ> processedIntersections = new List<XYZ>();

                            segmentIntersections = segmentIntersections.OrderBy(p => p.DistanceTo(start)).ToList();

                            bool[] used = new bool[segmentIntersections.Count];

                            for (int j = 0; j < segmentIntersections.Count - 1; j++)
                            {
                                if (used[j]) continue;

                                XYZ first = segmentIntersections[j];

                                for (int k = j + 1; k < segmentIntersections.Count; k++)
                                {
                                    if (used[k]) continue;

                                    XYZ second = segmentIntersections[k];
                                    double distance = first.DistanceTo(second);

                                    if (Math.Abs(distance - wallThickness) < wallThickness * 0.3)
                                    {
                                        XYZ middle = (first + second) / 2;
                                        processedIntersections.Add(middle);

                                        used[j] = true;
                                        used[k] = true;
                                        break;
                                    }
                                }
                            }

                            allIntersections.AddRange(processedIntersections);
                        }
                    }
                }
            }

            // Временное решение расчёта гибких элементов (2 из 2)
            if (linkInstance != null && allIntersections.Count > 0 && (cElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexPipeCurves ||
                cElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexDuctCurves))
            {
                List<XYZ> processedIntersections = new List<XYZ>();

                for (int i = 0; i < allIntersections.Count - 1; i += 2)
                {
                    XYZ first = allIntersections[i];
                    XYZ second = allIntersections[i + 1];

                    XYZ middle = (first + second) / 2;
                    processedIntersections.Add(middle);
                }

                allIntersections = processedIntersections;
            }

            // Корректировка по оси Z для АР (временное решение?)     
            if (department == "АР")
            {
                List<XYZ> allIntersectionsAR = new List<XYZ>();
                double adjustmentZ = 0;

                double offsetBottom = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                double offsetTop = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0;
                double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0;

                if (!wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).IsReadOnly)
                {
                    adjustmentZ = unconnectedHeight - offsetBottom + offsetTop;
                }

                foreach (var nsElement in allIntersections)
                {
                    XYZ adjusted = new XYZ(
                        nsElement.X,
                        nsElement.Y,
                        nsElement.Z + adjustmentZ
                    );

                    allIntersectionsAR.Add(adjusted);

                }

                allIntersections = allIntersectionsAR;
            }

            if (createdView) doc.Delete(view3D.Id);
            return allIntersections;
        }

        /// <summary>
        /// Получает трансформацию элемента, если он находится в связанном файле.
        /// </summary>
        private Autodesk.Revit.DB.Transform GetElementTransform(Document doc, Element element)
        {
            Document elementDoc = element.Document;
            if (doc.Equals(elementDoc))
                return Autodesk.Revit.DB.Transform.Identity;

            RevitLinkInstance linkInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .FirstOrDefault(rl => rl.GetLinkDocument()?.Equals(elementDoc) == true);

            return linkInstance?.GetTransform() ?? Autodesk.Revit.DB.Transform.Identity;
        }

        /// <summary>
        /// Стена. Метод для получения размеров стены
        /// </summary>
        double GetWallThickness(Element wall)
        {
            // Пробуем получить толщину через параметр WALL_ATTR_WIDTH_PARAM
            double wallThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? 0;

            if (wallThickness == 0 && wall is Wall actualWall)
            {
                WallType wallType = actualWall.Document.GetElement(actualWall.GetTypeId()) as WallType;
                if (wallType != null && wallType.GetCompoundStructure() != null)
                {
                    wallThickness = wallType.GetCompoundStructure().GetWidth();
                }
            }

            // Проверяем, находится ли стена в связанном файле
            Document wallDoc = wall.Document;
            RevitLinkInstance linkInstance = new FilteredElementCollector(wallDoc.Application.Documents.Cast<Document>()
                .FirstOrDefault(d => d.Title == wallDoc.Title))
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .FirstOrDefault(link => link.GetLinkDocument()?.Equals(wallDoc) == true);

            if (linkInstance != null)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc != null)
                {
                    WallType linkedWallType = linkedDoc.GetElement(wall.GetTypeId()) as WallType;
                    if (linkedWallType != null && linkedWallType.GetCompoundStructure() != null)
                    {
                        wallThickness = linkedWallType.GetCompoundStructure().GetWidth();
                    }
                }
            }

            return wallThickness;
        }

        /// <summary>
        /// XYZ. Обычная стена. Получаем пересечения для одного сегмента.
        /// </summary>
        private List<XYZ> GetIntersectionsForSegment(ReferenceIntersector intersector, XYZ start, XYZ end, Element wall, double wallThickness)
        {
            List<XYZ> intersections = new List<XYZ>();
            XYZ direction = (end - start).Normalize();
            double segmentLength = start.DistanceTo(end);

            IList<ReferenceWithContext> references = intersector.Find(start, direction);

            foreach (ReferenceWithContext refWithContext in references)
            {
                Reference reference = refWithContext.GetReference();
                XYZ intersectionPoint = reference.GlobalPoint;
                double distanceFromStart = start.DistanceTo(intersectionPoint);

                if (distanceFromStart > segmentLength)
                    continue;

                Wall wallElement = wall as Wall;
                Face wallFace = wallElement?.GetGeometryObjectFromReference(reference) as Face;

                if (wallFace == null)
                    continue;

                bool isCurvedWall = (wallElement.Location as LocationCurve)?.Curve is Arc;

                XYZ wallNormal;
                if (isCurvedWall)
                {
                    wallNormal = GetCurvedWallNormal(wallElement, intersectionPoint);
                }
                else
                {
                    UV uvPoint;
                    if (!TryGetUVFromFace(wallFace, intersectionPoint, out uvPoint))
                        continue;

                    wallNormal = wallFace.ComputeNormal(uvPoint).Normalize();
                }

                // Смещаем точку вдоль нормали стены
                XYZ shiftedPoint = intersectionPoint + wallNormal * wallThickness;

                IList<ReferenceWithContext> shiftedReferences = intersector.Find(shiftedPoint, direction);
                bool stillInWall = shiftedReferences.Any(r => r.GetReference().ElementId == wall.Id);

                if (stillInWall)
                {
                    intersections.Add(intersectionPoint);
                }
            }

            return intersections;
        }

        /// <summary>
        /// XYZ. Обычная стена. UV-координаты в рамках поверхности
        /// </summary>
        private bool TryGetUVFromFace(Face face, XYZ point, out UV uv)
        {
            uv = null;
            BoundingBoxUV bb = face.GetBoundingBox();

            for (double u = bb.Min.U; u <= bb.Max.U; u += 0.01)
            {
                for (double v = bb.Min.V; v <= bb.Max.V; v += 0.01)
                {
                    UV testUV = new UV(u, v);
                    XYZ testXYZ = face.Evaluate(testUV);

                    if (testXYZ.DistanceTo(point) < 0.01)
                    {
                        uv = testUV;
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// XYZ. Изогнутая стена. Вычисляем нормаль
        /// </summary>
        private XYZ GetCurvedWallNormal(Wall curvedWall, XYZ point)
        {
            LocationCurve locationCurve = curvedWall.Location as LocationCurve;
            if (locationCurve == null)
                return XYZ.Zero;

            Arc arc = locationCurve.Curve as Arc;
            if (arc == null)
                return XYZ.Zero;

            XYZ arcCenter = arc.Center;

            XYZ normal = (point - arcCenter).Normalize();

            return normal;
        }
    }
}