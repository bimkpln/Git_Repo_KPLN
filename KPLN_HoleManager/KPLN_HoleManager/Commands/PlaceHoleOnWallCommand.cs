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
    public class PlaceHoleOnWallCommand
    {
        private readonly string _userFullName;
        private readonly string _departmentName;
        private readonly Element _selectedWall;
      
        private string _departmentHoleName;
        private string _sendingDepartmentHoleName;
        private string _holeTypeName;

        private bool transactionStatus;
        private bool deleteLastHole;

        public PlaceHoleOnWallCommand(string userFullName, string departmentName, Element selectedElement, string departmentHoleName, string sendingDepartmentHoleName, string holeTypeName)
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
            var command = new PlaceHoleOnWallCommand(userFullName, departmentName, selectedElement, departmentHoleName, sendingDepartmentHoleName, holeTypeName);
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
                Wall wall = _selectedWall as Wall;
                LocationCurve locationCurve = wall.Location as LocationCurve;

                
                Element intersectingElement = null; // Выбор элемента

                while (true)
                {
                    try
                    {
                        // Фильтр, который разрешает выбирать только нужные элементы и линки
                        Reference selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, new CustomSelectionFilter(), "Выберите элемент (воздуховоды, трубы, лотки, короба или линк).");
                        Element firstSelectedElement = doc.GetElement(selectedRef);

                        if (firstSelectedElement is RevitLinkInstance revitLink)
                        {                         
                            Document linkedDoc = revitLink.GetLinkDocument();

                            if (linkedDoc == null)
                            {
                                TaskDialog.Show("Ошибка", "Не удалось получить связанный документ.");
                                break;
                            }

                            try
                            {
                                Reference linkedRef = uiDoc.Selection.PickObject(ObjectType.LinkedElement, "Выберите элемент внутри линка.");

                                if (linkedRef.LinkedElementId == null)
                                {
                                    TaskDialog.Show("Ошибка", "Не удалось получить элемент внутри линка.");
                                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                                    break;
                                }

                                Element linkedElement = linkedDoc.GetElement(linkedRef.LinkedElementId);

                                if (linkedElement != null)
                                {
                                    intersectingElement = linkedElement;
                                    break;
                                }
                            }
                            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                            {
                                TaskDialog.Show("Отмена", "Выбор элемента был отменён.");
                                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                                break;
                            }
                        }
                        else
                        {
                            intersectingElement = firstSelectedElement;
                            break;
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        TaskDialog.Show("Отмена", "Выбор элемента был отменён.");
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        break;
                    }

                    TaskDialog.Show("Предупреждение", "Выбранный элемент не является допустимым объектом. Операция отменена.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    break;
                }          

                if (intersectingElement == null)
                {
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                XYZ holeLocation = GetIntersectionPoint(doc, _selectedWall, intersectingElement, _departmentHoleName, _holeTypeName);

                if (holeLocation == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить точку пересечения.\nВыберите стену заново и повторите попытку.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
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
                using (Transaction tx = new Transaction(doc, $"KPLN. Разместить отверстие {familyFileName}"))
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

                    // Создаём отверстие без преднастроек (указывается только точка самого отверстия)
                    FamilyInstance holeInstance = doc.Create.NewFamilyInstance(holeLocation, holeSymbol, _selectedWall, StructuralType.NonStructural);

                    if (holeInstance == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Ошибка при создании отверстия.");
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        return;
                    }

                    // Задаём размеры отверстию, двигаем его и пишем данные в Extensible Storage
                    (double widthElement, double heightElement, double lengthElement) = GetElementSize(intersectingElement);
                    SetHoleDimensions(uiDoc, doc, tx, wall, intersectingElement, holeInstance, holeLocation, widthElement, heightElement);

                    if (deleteLastHole)
                    {
                        doc.Delete(holeInstance.Id);
                    }

                    if (transactionStatus == false)
                    {
                        tx.RollBack();
                        if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                        return;
                    }

                    // Обновление данных интерфейса
                    DockableManagerForm.Instance?.UpdateStatusCounts();
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                  
                    tx.Commit();                  
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
            }
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
        /// XYZ. Метод нахождения точки пересечения стены и входящего элемента.
        /// </summary>
        private XYZ GetIntersectionPoint(Document doc, Element wall, Element cElement, string department, string holeType)
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
            // Обработка ошибки вырезания отверстия в стене связного файла
            else if (linkInstance != null && department == "АР")
            {
                TaskDialog.Show("Ошибка", "Вы пытаетесь вырезать отверстие в стене из связанного файла. Операция отменена.");
                if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                return null;
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

            // Выбор точки в диалоговом окне
            if (allIntersections.Count == 0)
            {
                if (createdView) doc.Delete(view3D.Id);
                return null;
            }

            XYZ selectedIntersection;
            if (allIntersections.Count == 1)
            {
                selectedIntersection = allIntersections[0]; 
            }
            else
            {               
                selectedIntersection = ShowIntersectionSelectionDialog(allIntersections);

                if (selectedIntersection == null)
                {
                    if (createdView) doc.Delete(view3D.Id);
                    return null;
                }
            }

            // Корректировка по оси Z для АР (временное решение?)
            double adjustmentZ = 0;
            if (department == "АР")
            {
                double offsetBottom = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                double offsetTop = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0;
                double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0;

                if (!wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).IsReadOnly)
                {
                    adjustmentZ = unconnectedHeight - offsetBottom + offsetTop;
                }
            }

            selectedIntersection = new XYZ(selectedIntersection.X, selectedIntersection.Y, selectedIntersection.Z + adjustmentZ);

            if (createdView) doc.Delete(view3D.Id);

            return selectedIntersection;
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

        /// <summary>
        /// XYZ. Отображаем диалоговое окно для выбора точки пересечения.
        /// </summary>
        private XYZ ShowIntersectionSelectionDialog(List<XYZ> intersections)
        {
            if (intersections == null || intersections.Count == 0)
            {
                TaskDialog.Show("Ошибка", "Не найдено точек пересечения.");
                return null;
            }

            List<string> intersectionStrings = new List<string>();
            for (int i = 0; i < intersections.Count; i++)
            {
                intersectionStrings.Add($"Точка {i + 1}: X = {intersections[i].X:F2}, Y = {intersections[i].Y:F2}, Z = {intersections[i].Z:F2}");
            }

            TaskDialog td = new TaskDialog("Выберите точку пересечения");
            td.MainInstruction = "Выберите одну из точек пересечения:";

            for (int i = 0; i < intersectionStrings.Count; i++)
            {
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1 + i, intersectionStrings[i]);
            }

            TaskDialogResult result = td.Show();

            for (int i = 0; i < intersectionStrings.Count; i++)
            {
                if (result == (TaskDialogResult)(TaskDialogCommandLinkId.CommandLink1 + i))
                {
                    return intersections[i];
                }
            }

            return null;
        }

        /// <summary>
        /// Пересекающий стену элемент. Метод для получения размеров элемента системы
        /// </summary>
        private (double width, double height, double length) GetElementSize(Element element)
        {
            double width = 0, height = 0, length = 0;

            Category category = element.Category;
            if (category == null) return (0, 0, 0);

            // Воздуховоды и гибкие воздуховоды
            if (category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves
                || category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexDuctCurves)
            {
                double diameter = GetParameterValue(element, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (diameter > 0)
                {
                    width = height = diameter;
                }
                else
                {
                    height = GetParameterValue(element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                    width = GetParameterValue(element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                }
            }
            // Трубы и гибкие трубы
            else if (category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves
                || category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexPipeCurves)
            {
                double diameter = GetParameterValue(element, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (diameter > 0)
                {
                    width = height = diameter;
                }
            }
            // Кабельные лотки
            else if (category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
            {
                height = GetParameterValue(element, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                width = GetParameterValue(element, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            }
            // Короба
            else if (category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit)
            {
                double diameter = GetParameterValue(element, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (diameter > 0)
                {
                    width = height = diameter;
                }
            }

            // Определяем длину элемента (универсально для всех категорий)
            length = GetParameterValue(element, BuiltInParameter.CURVE_ELEM_LENGTH);

            return (width, height, length);
        }

        /// <summary>
        /// Пересекающий стену элемент. Метод для получения значения параметра с переводом футов в метры
        /// </summary>
        private double GetParameterValue(Element element, BuiltInParameter paramId)
        {
            Parameter param = element.get_Parameter(paramId);
            if (param != null && param.HasValue)
            {
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
            }
            return 0;
        }

        /// <summary>
        /// Отверстие. Метод изменения размера отверстия и его разворот после создания XYZ
        /// </summary>
        public void SetHoleDimensions(UIDocument uiDoc, Document doc, Transaction tx, 
            Wall wall, Element intersectingElement, FamilyInstance holeInstance, XYZ holeLocation, double widthElement, double heightElement)
        {
            deleteLastHole = false;

            double offset = -900; // Заведомо нереалистичное смещение
            List<string> settings = DockableManagerFormSettings.LoadSettings();
           
            if (settings != null && settings[5] != "Не выбрано")
            {
                if (!double.TryParse(settings[5], out offset))
                {
                    transactionStatus = false;
                    TaskDialog.Show("Ошибка", "Некорректное значение отступа в настройках.");
                    return;
                }
            }
            else
            {
                Forms.sChoiseHoleIndentation sChoiseHoleIndentation = new Forms.sChoiseHoleIndentation();

                if (sChoiseHoleIndentation.ShowDialog() == true)
                {
                    offset = sChoiseHoleIndentation.SelectedOffset;
                }
                else
                {
                    transactionStatus = false;
                    TaskDialog.Show("Внимание", "Операция отменена пользователем.");
                    return;
                }
            }

            // Обработка в случае ошибки поиска смещения
            if (offset == -900)
            {
                transactionStatus = false;
                TaskDialog.Show("Ошибка", "Не удалось обработать смещение от отверстия. Операция была отменена.");
                return;
            }

            Parameter widthParam = holeInstance.LookupParameter("Ширина");
            Parameter heightParam = holeInstance.LookupParameter("Высота");
            Parameter heightOtherParam = holeInstance.LookupParameter("КП_Р_Высота");
            Parameter diametrParam = holeInstance.LookupParameter("Диаметр");

            double internalHeight = UnitUtils.ConvertToInternalUnits(heightElement + offset, UnitTypeId.Millimeters);
            double internalWidth = UnitUtils.ConvertToInternalUnits(widthElement + offset, UnitTypeId.Millimeters);
            double internalDiametr = UnitUtils.ConvertToInternalUnits(Math.Max(widthElement, heightElement) + offset, UnitTypeId.Millimeters);

            Curve wallCurve = (wall.Location as LocationCurve).Curve;
            XYZ wallDirection;

            if (wallCurve is Line line)
            {
                wallDirection = line.Direction;
            }
            else if (wallCurve is Arc arc)
            {
                XYZ arcCenter = arc.Center;

                XYZ radialVector = (holeLocation - arcCenter).Normalize();

                wallDirection = new XYZ(-radialVector.Y, radialVector.X, 0);
            }
            else
            {
                return;
            }
            XYZ wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize(); ;

            double angle = 0;

            if (_departmentHoleName == "АР")
            {
                wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();

                if (_holeTypeName == "SquareHole")
                {
                    heightParam.Set(internalHeight);
                    widthParam.Set(internalWidth);
                }
                else
                {
                    heightOtherParam.Set(internalDiametr);
                }

                LocationPoint holeLocationAR = holeInstance.Location as LocationPoint;
                if (holeLocationAR != null)
                {
                    if (_holeTypeName == "SquareHole")
                    {
                        double shiftZ = -0.5 * internalHeight;
                        XYZ moveVector = new XYZ(0, 0, shiftZ);
                        ElementTransformUtils.MoveElement(holeInstance.Document, holeInstance.Id, moveVector);
                    }
                    else
                    {
                        double shiftZ = -0.5 * Math.Max(internalWidth, internalHeight);
                        XYZ moveVector = new XYZ(0, 0, shiftZ);
                        ElementTransformUtils.MoveElement(holeInstance.Document, holeInstance.Id, moveVector);
                    }
                }
            }
            else if (_departmentHoleName == "ОВиК" || _departmentHoleName == "ВК"
                || _departmentHoleName == "ЭОМ" || _departmentHoleName == "СС")
            {
                if (_holeTypeName == "SquareHole")
                {
                    heightParam.Set(internalHeight);
                    widthParam.Set(internalWidth);
                }
                else
                {
                    diametrParam.Set(internalDiametr);
                }
       
                if (!wall.Document.IsLinked)
                {
                    angle = Math.Atan2(wallDirection.Y, wallDirection.X);
                    Line rotationAxis = Line.CreateBound(holeLocation, holeLocation + XYZ.BasisZ);

                    ElementTransformUtils.RotateElement(wall.Document, holeInstance.Id, rotationAxis, angle);
                    double wallThickness = GetWallThickness(wall);

                    // Обработка ИОСных отверстий
                    if (wallThickness > 0)
                    {                       
                        // Определяем сторону, в которую нужно сдвигать (в сторону intersectingElement)
                        XYZ intersectingElementLocation = (intersectingElement.Location as LocationPoint)?.Point ?? XYZ.Zero;
                        XYZ directionToIntersecting = (intersectingElementLocation - holeLocation).Normalize();

                        // Проверяем, в каком направлении двигаться - в сторону нормали или против
                        double dotProduct = wallNormal.DotProduct(directionToIntersecting);
                        XYZ moveDirection = dotProduct >= 0 ? wallNormal : -wallNormal;

                        XYZ moveOffset = moveDirection * (wallThickness / 2);

                        // Проверяем, выйдет ли отверстие за пределы стены после смещения
                        XYZ newHoleLocation = holeLocation + moveOffset;

                        if (!IsPointInsideWall(newHoleLocation, wall, wallThickness))
                        {
                            moveOffset = -moveOffset;
                            newHoleLocation = holeLocation + moveOffset;
                        }
                        if (IsPointInsideWall(newHoleLocation, wall, wallThickness))
                        {
                            ElementTransformUtils.MoveElement(wall.Document, holeInstance.Id, moveOffset);
                        }
                    }
                }
                else
                {
                    Autodesk.Revit.DB.Transform transform = GetElementTransform(doc, wall);
                    XYZ transformedWallDirection = transform.OfVector(wallDirection);

                    angle = Math.Atan2(transformedWallDirection.Y, transformedWallDirection.X);
                    Line rotationAxis = Line.CreateBound(holeLocation, holeLocation + XYZ.BasisZ);

                    if (!holeInstance.Document.IsLinked)
                    {
                        ElementTransformUtils.RotateElement(holeInstance.Document, holeInstance.Id, rotationAxis, angle);
                    }
                }
            }

            string wallIdString = wall.Id.IntegerValue.ToString();
            string intersectingElementIdString = intersectingElement.Id.IntegerValue.ToString();

            // Проверка пересечений
            List<FamilyInstance> intersectingHoles = GetIntersectingHoles(doc, holeInstance);

            if (intersectingHoles.Any())
            {
                deleteLastHole = true;
                List<XYZ> holeCenters = new List<XYZ>();

                // Центр каждого отверстия + вычесление общего центра
                foreach (FamilyInstance hole in intersectingHoles)
                {
                    XYZ center = null;

                    BoundingBoxXYZ bbox = hole.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        center = (bbox.Min + bbox.Max) / 2;
                    }                  

                    if (center != null)
                    {
                        holeCenters.Add(center);
                    }
                }

                BoundingBoxXYZ bboxHoleInstance = holeInstance.get_BoundingBox(null);
                if (bboxHoleInstance != null)
                {
                    XYZ holeInstanceCenter = (bboxHoleInstance.Min + bboxHoleInstance.Max) / 2;
                    holeCenters.Add(holeInstanceCenter); 
                }

                double avgX = holeCenters.Average(p => p.X);
                double avgY = holeCenters.Average(p => p.Y);
                double avgZ = holeCenters.Average(p => p.Z);
                XYZ newHoleLocation = new XYZ(avgX, avgY, avgZ);

                // Корректировка по оси Z для АР (временное решение?)
                if (_departmentHoleName == "АР")
                {
                    double offsetBottom = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                    double offsetTop = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0;
                    double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0;

                    double adjustmentZ = 0;

                    if (!wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).IsReadOnly)
                    {
                        adjustmentZ = unconnectedHeight - offsetBottom + offsetTop;
                    }

                    newHoleLocation = new XYZ(newHoleLocation.X, newHoleLocation.Y, newHoleLocation.Z + adjustmentZ);
                }

                // Загрузка семейства
                string newFamilyFileName = GetFamilyFileName(_departmentHoleName, "SquareHole");
                if (string.IsNullOrEmpty(newFamilyFileName))
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить файл семейства.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                string newFamilyPath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Families", newFamilyFileName);

                if (!File.Exists(newFamilyPath))
                {
                    TaskDialog.Show("Ошибка", $"Файл семейства не найден:\n{newFamilyPath}");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                Family newFamily = LoadOrGetExistingFamily(doc, newFamilyPath);

                if (newFamily == null)
                {
                    tx.RollBack();
                    TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                // Получаем FamilySymbol
                FamilySymbol newHoleSymbol = GetFamilySymbol(newFamily, doc);
                if (newHoleSymbol == null)
                {
                    tx.RollBack();
                    TaskDialog.Show("Ошибка", "Не удалось получить типоразмер семейства.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                if (!newHoleSymbol.IsActive)
                {
                    newHoleSymbol.Activate();
                    doc.Regenerate();
                }

                // Создаём отверстие без преднастроек (указывается только точка самого отверстия)
                FamilyInstance newHoleInstance = doc.Create.NewFamilyInstance(newHoleLocation, newHoleSymbol, _selectedWall, StructuralType.NonStructural);

                if (newHoleInstance == null)
                {
                    tx.RollBack();
                    TaskDialog.Show("Ошибка", "Ошибка при создании отверстия.");
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                Parameter newWidthParam = newHoleInstance.LookupParameter("Ширина");
                Parameter newHightParam = newHoleInstance.LookupParameter("Высота");

                double minX = double.MaxValue, maxX = double.MinValue;
                double minY = double.MaxValue, maxY = double.MinValue;
                double minZ = double.MaxValue, maxZ = double.MinValue;

                foreach (FamilyInstance hole in intersectingHoles)
                {
                    BoundingBoxXYZ bbox = hole.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        minX = Math.Min(minX, bbox.Min.X);
                        maxX = Math.Max(maxX, bbox.Max.X);

                        minY = Math.Min(minY, bbox.Min.Y);
                        maxY = Math.Max(maxY, bbox.Max.Y);

                        minZ = Math.Min(minZ, bbox.Min.Z);
                        maxZ = Math.Max(maxZ, bbox.Max.Z);
                    }
                }

                BoundingBoxXYZ bboxHoleInstanceNew = holeInstance.get_BoundingBox(null);
                if (bboxHoleInstanceNew != null)
                {
                    minX = Math.Min(minX, bboxHoleInstance.Min.X);
                    maxX = Math.Max(maxX, bboxHoleInstance.Max.X);

                    minY = Math.Min(minY, bboxHoleInstance.Min.Y);
                    maxY = Math.Max(maxY, bboxHoleInstance.Max.Y);

                    minZ = Math.Min(minZ, bboxHoleInstance.Min.Z);
                    maxZ = Math.Max(maxZ, bboxHoleInstance.Max.Z);
                }

                // Проверяем ориентацию стены
                bool isWallVertical = Math.Abs(wallNormal.Z) > 0.9; 
                double width = isWallVertical ? maxX - minX : maxY - minY; 
                double height = maxZ - minZ; 
      
                newHightParam.Set(height);
                newWidthParam.Set(width);

                // Вращаем отверстие
                if (_departmentHoleName != "АР")
                {
                    Line rotationAxis = Line.CreateBound(newHoleLocation, newHoleLocation + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(newHoleInstance.Document, newHoleInstance.Id, rotationAxis, angle);
                }
                else
                {               
                    double shiftZ = -0.5 * height;
                    XYZ moveVector = new XYZ(0, 0, shiftZ);
                    ElementTransformUtils.MoveElement(newHoleInstance.Document, newHoleInstance.Id, moveVector);
                }

                // Генерация данных интерфейса и удаление отверстий
                string allIntersectingHoleIdString = $"{holeInstance.Id}, " + string.Join(", ", intersectingHoles.Select(h => h.Id.IntegerValue));
                string allIntersectingElementIdString = intersectingElementIdString;

                foreach (FamilyInstance intersectingHole in intersectingHoles)
                {
                    List<List<string>> holeMessages = _iDataProcessor.GetHoleTaskMessages(doc, intersectingHole.Id.IntegerValue.ToString());

                    if (holeMessages.Any())
                    {
                        List<string> lastMessage = holeMessages.Last();
                        if (lastMessage.Count > 8)
                        {
                            allIntersectingElementIdString += $", {lastMessage[8]}";
                        }
                    }

                    doc.Delete(intersectingHole.Id);
                }

                // Запись данных в хранилище             
                ExtensibleStorageHelper.AddChatMessage(
                    newHoleInstance,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    _userFullName,
                    _departmentName,
                    _departmentHoleName,
                    _sendingDepartmentHoleName,
                    wallIdString,
                    allIntersectingElementIdString,
                    "Без статуса",
                    "00",
                    $"Отверстия объеденены: {allIntersectingHoleIdString}"
                );

                if (transactionStatus == false)
                {
                    tx.RollBack();
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                CreateHoleTaskIntrerface(uiDoc, doc, newHoleInstance, _userFullName, _departmentName, _departmentHoleName, _sendingDepartmentHoleName, wallIdString, allIntersectingElementIdString);
            }

            else 
            {
                // Запись данных в хранилище             
                ExtensibleStorageHelper.AddChatMessage(
                    holeInstance,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    _userFullName,
                    _departmentName,
                    _departmentHoleName,
                    _sendingDepartmentHoleName,
                    wallIdString,
                    intersectingElementIdString,
                    "Без статуса",
                    "00",
                    "Отверстие создано"
                );

                if (transactionStatus == false)
                {
                    tx.RollBack();
                    if (DockableManagerForm.Instance != null) DockableManagerForm.Instance.IsEnabled = true;
                    return;
                }

                CreateHoleTaskIntrerface(uiDoc, doc, holeInstance, _userFullName, _departmentName, _departmentHoleName, _sendingDepartmentHoleName, wallIdString, intersectingElementIdString);
            }         
        }

        /// <summary>
        /// Отверстие. Вспомогательный метод к SetHoleDimensions для проверки, остаётся ли точка внутри границ стены
        /// </summary>
        private bool IsPointInsideWall(XYZ point, Element wall, double wallThickness)
        {
            BoundingBoxXYZ bbox = wall.get_BoundingBox(null);
            return (point.X >= bbox.Min.X && point.X <= bbox.Max.X) &&
                   (point.Y >= bbox.Min.Y && point.Y <= bbox.Max.Y) &&
                   (point.Z >= bbox.Min.Z && point.Z <= bbox.Max.Z);
        }

        /// <summary>
        /// Отверстие. Вспомогательный метод к SetHoleDimensions поиск пересекающихся отверстий
        /// </summary>
        private static List<FamilyInstance> GetIntersectingHoles(Document doc, FamilyInstance holeInstance)
        {
            List<FamilyInstance> intersectingHoles = new List<FamilyInstance>();
            string holeFamilyName = holeInstance.Symbol.Family.Name;

            // Определяем, с какими отверстиями искать пересечения
            List<string> targetHoleFamilies = new List<string>();

            if (holeFamilyName == "199_Отверстие прямоугольное_(Об_Стена)" || holeFamilyName == "199_Отверстие круглое_(Об_Стена)")
            {
                targetHoleFamilies.Add("199_Отверстие прямоугольное_(Об_Стена)");
                targetHoleFamilies.Add("199_Отверстие круглое_(Об_Стена)");
            }
            else if (holeFamilyName == "501_ЗИ_Отверстие_Прямоугольное_Стена_(Об)" || holeFamilyName == "501_ЗИ_Отверстие_Круглое_Стена_(Об)")
            {
                targetHoleFamilies.Add("501_ЗИ_Отверстие_Прямоугольное_Стена_(Об)");
                targetHoleFamilies.Add("501_ЗИ_Отверстие_Круглое_Стена_(Об)");
            }
            else
            {
                return intersectingHoles;
            }

            BoundingBoxXYZ bbox = holeInstance.get_BoundingBox(null);
            if (bbox == null) return intersectingHoles;

            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(new Outline(bbox.Min, bbox.Max));

            List<FamilyInstance> candidateInstances = new FilteredElementCollector(doc)
                .WherePasses(filter)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Id != holeInstance.Id && targetHoleFamilies.Contains(fi.Symbol.Family.Name)) 
                .ToList();

            Solid holeSolid = GetSolidFromInstance(holeInstance);
            if (holeSolid == null) return intersectingHoles;

            foreach (FamilyInstance otherInstance in candidateInstances)
            {
                Solid otherSolid = GetSolidFromInstance(otherInstance);
                if (otherSolid == null) continue;

                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(holeSolid, otherSolid, BooleanOperationsType.Intersect);
                if (intersection != null && intersection.Volume > 0)
                {
                    intersectingHoles.Add(otherInstance);
                }
            }

            return intersectingHoles;
        }

        /// <summary>
        /// Отверстие. Вспомогательный метод к GetIntersectingHoles прасчёт геометрии пересекающихся элементов
        /// </summary>
        private static Solid GetSolidFromInstance(FamilyInstance instance)
        {
            Options options = new Options { ComputeReferences = false };
            GeometryElement geometry = instance.get_Geometry(options);

            foreach (GeometryObject obj in geometry)
            {
                if (obj is GeometryInstance geomInstance)
                {
                    Autodesk.Revit.DB.Transform transform = geomInstance.Transform;
                    foreach (GeometryObject instanceObj in geomInstance.GetSymbolGeometry())
                    {
                        if (instanceObj is Solid solid && solid.Volume > 0)
                        {
                            return SolidUtils.CreateTransformed(solid, transform); 
                        }
                    }
                }
                else if (obj is Solid solid && solid.Volume > 0)
                {
                    return solid;
                }
            }
            return null;
        }

        /// <summary>
        /// Отверстие. Вспомогательный метод отрисовки интерфейса задания
        /// </summary>
        private void CreateHoleTaskIntrerface(UIDocument uiDoc, Document doc, FamilyInstance holeInstance, 
            string userFullName, string departmentName, string departamentFrom, string departamentIn, string wallID, string sEllementID)
        {
            string holeID = holeInstance.Id.IntegerValue.ToString();
            string holeName = $"{holeInstance.Symbol.Family.Name} - {holeInstance.Symbol.Name}";

            DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Общая информация
            DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Панель действия
            DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Чат-область
            DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Поле добавления сообщения

            ///////////////////// Блок 1. Базовая информация
            TextBlock generalInfoTextBlock = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Padding = new Thickness(8)
            };

            // Заполение базовой информации
            generalInfoTextBlock.Inlines.Add(new Run(holeName) { FontWeight = FontWeights.Bold });
            generalInfoTextBlock.Inlines.Add(new Run($" ({holeID})\n"));
            generalInfoTextBlock.Inlines.Add(new Run("Без статуса") { FontWeight = FontWeights.Bold, Foreground = Brushes.Gray });
            generalInfoTextBlock.Inlines.Add(new Run($"\n"));
            generalInfoTextBlock.Inlines.Add(new Run("Стена: ") { FontWeight = FontWeights.Bold });
            generalInfoTextBlock.Inlines.Add(new Run($"{wallID}\n"));
            generalInfoTextBlock.Inlines.Add(new Run("Элементы в отверстии: ") { FontWeight = FontWeights.Bold });
            generalInfoTextBlock.Inlines.Add(new Run($"{sEllementID}"));

            System.Windows.Controls.Grid.SetRow(generalInfoTextBlock, 0);
            DockableManagerForm.Instance.InfoHolePanel.Children.Add(generalInfoTextBlock);

            ///////////////////// Блок 2. Изменение статуса отверстия
            StackPanel decisionPanel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Margin = new Thickness(2, 0, 0, 0),
            };

            // Кнопка "Да"
            System.Windows.Controls.Button yesButton = new System.Windows.Controls.Button
            {
                Content = "✔️",
                Width = 30,
                Height = 30,
                Background = Brushes.Green,
                Foreground = Brushes.White,
                Margin = new Thickness(6, 0, 15, 0)
            };
           
            yesButton.Click += (si, evi) =>
            {
                if (!int.TryParse(holeID, out int holeElementId))
                {
                    return;
                }

                ExtensibleStorageHelper.AddChatMessage(
                        holeInstance,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        userFullName,
                        departmentName,
                        departamentFrom,
                        departamentIn,
                        wallID,
                        sEllementID,
                        "Подтверждение",
                        "10",
                        $"Смена статуса на отверстия `Подтверждение`"
                        );

                DockableManagerForm.Instance.InfoHolePanel.Children.Clear();
                DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Clear();
                DockableManagerForm.Instance.UpdateStatusCounts();

                TaskDialog.Show("Информация", $"{holeName} ({holeID}). Статус обновлён на ``Подтверждение``");
            };
       
            TextBlock taskITextBlock = new TextBlock
            {
                Text = $"{departamentFrom}",
                Width = 55,
                Background = Brushes.LightGray,
                Padding = new Thickness(7),
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 4, 0)
            };

            TextBlock taskOTextBlock = new TextBlock
            {
                Text = $"{departamentIn}",
                Width = 55,
                Background = Brushes.LightGray,
                Padding = new Thickness(7),
                TextAlignment = TextAlignment.Center
            };

            // Добавляем элементы в панель
            if (departmentName != "BIM")
            {
                decisionPanel.Children.Add(yesButton);                        
            }
            else
            {
                decisionPanel.Margin = new Thickness(8, 0, 0, 0);
            }

            decisionPanel.Children.Add(taskITextBlock);
            decisionPanel.Children.Add(taskOTextBlock);

            System.Windows.Controls.Grid.SetRow(decisionPanel, 1);
            DockableManagerForm.Instance.InfoHolePanel.Children.Add(decisionPanel);

            ///////////////////// Блок 3. Чат и ScrollViewer
            ScrollViewer messagesScrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Margin = new Thickness(5, 8, 5, 8)
            };

            StackPanel messagesPanel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Stretch
            };

            // Переменная для хранения координат предыдущего сообщения
            string previousCoordinates = null;

            foreach (List<string> fullHoleInfoParts in _iDataProcessor.GetHoleTaskMessages(doc, holeID))
            {
                if (fullHoleInfoParts.Count > 10)
                {
                    string mDate = fullHoleInfoParts[0];
                    string mName = fullHoleInfoParts[1];
                    string mDepartmentFrom = fullHoleInfoParts[3];
                    string mDepartmentTo = fullHoleInfoParts[4];
                    string mCoordinates = fullHoleInfoParts[9];
                    string mMessageText = fullHoleInfoParts[12];

                    TextBlock messageTextBlock = new TextBlock { TextWrapping = TextWrapping.Wrap };

                    messageTextBlock.Inlines.Add(new Run($"{mDate}") { FontWeight = FontWeights.Bold, Foreground = Brushes.BlueViolet });
                    messageTextBlock.Inlines.Add(new Run($" | {mName} ({mDepartmentFrom} → {mDepartmentTo})\n"));

                    messageTextBlock.Inlines.Add(new Run("Координаты: ") { FontWeight = FontWeights.Bold });

                    // Проверяем, изменились ли координаты с прошлого сообщения
                    Brush coordinatesColor = previousCoordinates != null && previousCoordinates != mCoordinates
                        ? Brushes.Red
                        : Brushes.Black;

                    messageTextBlock.Inlines.Add(new Run($"{mCoordinates}\n") { Foreground = coordinatesColor });

                    messageTextBlock.Inlines.Add(new Run("💬  Сообщение: ") { FontWeight = FontWeights.Bold });
                    messageTextBlock.Inlines.Add(new Run($"{mMessageText}"));

                    Border messageBorder = new Border
                    {
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 245, 255)),
                        BorderBrush = Brushes.LightGray,
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(5),
                        Padding = new Thickness(8),
                        Child = messageTextBlock
                    };

                    // Обновляем previousCoordinates для следующей итерации
                    previousCoordinates = mCoordinates;

                    messagesPanel.Children.Add(messageBorder);
                }
            }

            messagesScrollViewer.Content = messagesPanel;
            System.Windows.Controls.Grid.SetRow(messagesScrollViewer, 2);
            DockableManagerForm.Instance.InfoHolePanel.Children.Add(messagesScrollViewer);

            ///////////////////// Блок 4. Блок добавления комментария
            System.Windows.Controls.Grid sendMessagesPanel = new System.Windows.Controls.Grid
            {
                Margin = new Thickness(5, 5, 5, 5),
                VerticalAlignment = VerticalAlignment.Bottom
            };

            sendMessagesPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            sendMessagesPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });

            System.Windows.Controls.TextBox commentTextBox = new System.Windows.Controls.TextBox
            {
                Height = 45,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(5, 5, 5, 5),
                Margin = new Thickness(5, 0, 3, 0),
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            System.Windows.Controls.Grid.SetColumn(commentTextBox, 0);

            System.Windows.Controls.Button sendButton = new System.Windows.Controls.Button
            {
                Content = "🔼",
                Height = 45,
                Width = 30,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                IsEnabled = false
            };

            commentTextBox.TextChanged += (si, evi) =>
            {
                sendButton.IsEnabled = !string.IsNullOrWhiteSpace(commentTextBox.Text);
            };

            sendButton.Click += (si, evi) =>
            {
                if (!int.TryParse(holeID, out int holeElementId))
                {
                    return;
                }

                string commentText = commentTextBox.Text;

                ExtensibleStorageHelper.AddChatMessage(
                    holeInstance,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    userFullName,
                    departmentName,
                    departamentFrom,
                    departamentIn,
                    wallID,
                    sEllementID,
                    "Без статуса",
                    "00",
                    commentText
                    );

                DockableManagerForm.Instance.InfoHolePanel.Children.Clear();
                DockableManagerForm.Instance.InfoHolePanel.RowDefinitions.Clear();

                CreateHoleTaskIntrerface(uiDoc, doc, holeInstance, userFullName, departmentName, departamentFrom, departamentIn, wallID, sEllementID);
            };

            System.Windows.Controls.Grid.SetColumn(sendButton, 1);

            sendMessagesPanel.Children.Add(commentTextBox);
            sendMessagesPanel.Children.Add(sendButton);

            System.Windows.Controls.Grid.SetRow(sendMessagesPanel, 3);
            DockableManagerForm.Instance.InfoHolePanel.Children.Add(sendMessagesPanel);
        }
    }

    // Фильтр на выбор пересекающего стену элемента
    public class CustomSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem == null) return false;

            return elem is RevitLinkInstance || CategoryFilter.IsAllowedElement(elem);
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }

    // Фильтр на элементы системы
    public static class CategoryFilter
    {
        public static bool IsAllowedElement(Element element)
        {
            if (element == null) return false;
          
            var allowedCategories = new HashSet<int>
            {
            (int)BuiltInCategory.OST_DuctCurves,
            (int)BuiltInCategory.OST_FlexDuctCurves,
            (int)BuiltInCategory.OST_PipeCurves,
            (int)BuiltInCategory.OST_FlexPipeCurves,
            (int)BuiltInCategory.OST_CableTray,
            (int)BuiltInCategory.OST_Conduit
            };

            return element.Category != null && allowedCategories.Contains(element.Category.Id.IntegerValue);
        }
    }

    // Перезаписываем параметры вложенных семейств
    public class ReplaceNestedFamilyHandler : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true; 
            return true; 
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true; 
        }
    }
}