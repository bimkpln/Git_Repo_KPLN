﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_HoleManager.Forms;

namespace KPLN_HoleManager.Commands
{
    public class PlaceHoleOnWallCommand
    {
        private readonly string _userFullName;
        private readonly string _departmentName;
        private readonly Element _selectedWall;

        private readonly string _departmentHoleName;
        private readonly string _sendingDepartmentHoleName;
        private readonly string _holeTypeName;

        private bool transactionStatus;

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

            try
            {
                // Стена и её кривая
                Wall wall = _selectedWall as Wall;
                LocationCurve locationCurve = wall.Location as LocationCurve;

                // Список допустимых категорий элементов
                BuiltInCategory[] allowedCategories = new BuiltInCategory[]
                {
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit
                };

                // Выбираемый элемент, который пересекает стену
                Element intersectingElement = null;

                while (true)
                {
                    try
                    {
                        TaskDialog.Show("Выбор элемента", "Для продолжения работы выберите элемент, который пересекается с выбранной стеной.");
                        Reference selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Выберите элемент, который пересекается со стеной.");
                        intersectingElement = doc.GetElement(selectedRef);

                        if (intersectingElement == null)
                        {
                            TaskDialog.Show("Ошибка", "Не удалось выбрать элемент для пересечения. Попробуйте снова.");
                            continue;
                        }

                        if (!allowedCategories.Contains((BuiltInCategory)intersectingElement.Category.Id.IntegerValue))
                        {
                            TaskDialog.Show("Ошибка", "Выбранный элемент не является допустимым объектом (воздуховоды, трубы, кабельные лотки или короба). Попробуйте снова.");
                            continue;
                        }
          
                        break;                       
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        TaskDialog.Show("Отмена", "Выбор элемента был отменён.");
                        return; 
                    }
                }
            
                // Определяем точку пересечения
                XYZ holeLocation = GetIntersectionPoint(doc, _selectedWall, intersectingElement, _departmentHoleName, _holeTypeName);

                if (holeLocation == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить точку пересечения.\nВыберите стену заново и повторите попытку.");
                    return;
                }

                // Определение файла семейства
                string familyFileName = GetFamilyFileName(_departmentHoleName, _holeTypeName);

                if (string.IsNullOrEmpty(familyFileName))
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить файл семейства.");
                    return;
                }

                string pluginFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string familyPath = Path.Combine(pluginFolder, "Families", familyFileName);

                if (!File.Exists(familyPath))
                {
                    TaskDialog.Show("Ошибка", $"Файл семейства не найден:\n{familyPath}");
                    return;
                }

                // Запуск общей транзакции
                using (Transaction tx = new Transaction(doc, $"KPLN. Разместить отверстие {familyFileName}"))
                {
                    tx.Start();

                    transactionStatus = true;

                    // Загружаем семейство (если оно уже загружено, берём существующее)
                    Family family = LoadOrGetExistingFamily(doc, familyPath);

                    if (family == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                        return;
                    }

                    // Получаем FamilySymbol
                    FamilySymbol holeSymbol = GetFamilySymbol(family, doc);
                    if (holeSymbol == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Не удалось получить типоразмер семейства.");
                        return;
                    }

                    // Активируем типоразмер, если он не активен
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
                        return;
                    }

                    // Задаём размеры отверстию, двигаем его и пишем данные в Extensible Storage
                    (double widthElement, double heightElement, double lengthElement) = GetElementSize(intersectingElement);

                    SetHoleDimensions(tx, wall, intersectingElement, holeInstance, holeLocation, widthElement, heightElement);

                    // Если всё же что-то поёдт не так - отменяем транзакцию :)
                    if (transactionStatus == false)
                    {
                        tx.RollBack();
                        return;
                    }

                    // Обновление данных интерфейса
                    DockableManagerForm.Instance?.UpdateStatusCounts();

                    tx.Commit();                  
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Семейство. Метод выбора файла семейства
        /// </summary>
        private string GetFamilyFileName(string department, string holeType)
        {
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
            // Получаем геометрию элемента (трубы, воздуховода и т.д.)
            LocationCurve cElementLocation = cElement.Location as LocationCurve;
            if (cElementLocation == null)
                return null;

            Curve cElementCurve = cElementLocation.Curve;
            List<XYZ> segmentPoints = cElementCurve.Tessellate() as List<XYZ>;

            if (segmentPoints == null || segmentPoints.Count < 2)
                return null;

            // Получаем или создаем 3D-вид
            View3D view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);

            bool createdView = false;

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
                return null;
            }
                    
            // Создаем ReferenceIntersector для стены
            ReferenceIntersector intersector = new ReferenceIntersector(wall.Id, FindReferenceTarget.Face, view3D);
            List<XYZ> allIntersections = new List<XYZ>();

            // Получаем толщину стены
            double wallThickness = GetWallThickness(wall);

            // Проходим по сегментам кривой
            for (int i = 0; i < segmentPoints.Count - 1; i++)
            {
                XYZ start = segmentPoints[i];
                XYZ end = segmentPoints[i + 1];

                // Получаем пересечения для текущего сегмента
                List<XYZ> segmentIntersections = GetIntersectionsForSegment(intersector, start, end, wall, wallThickness);
                allIntersections.AddRange(segmentIntersections);
            }

            if (allIntersections.Count == 0)
            {
                if (createdView) doc.Delete(view3D.Id);
                return null;
            }

            // Определяем точку пересечения
            XYZ selectedIntersection;
            if (allIntersections.Count == 1)
            {
                selectedIntersection = allIntersections[0]; // Автоматический выбор, если точка одна
            }
            else
            {
                // Показываем диалог, если точек несколько
                selectedIntersection = ShowIntersectionSelectionDialog(allIntersections);

                if (selectedIntersection == null)
                {
                    if (createdView) doc.Delete(view3D.Id);
                    return null;
                }
            }

            // Корректировка по оси Z для АР
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
        /// XYZ. Получаем пересечения для одного сегмента.
        /// </summary>
        private List<XYZ> GetIntersectionsForSegment(ReferenceIntersector intersector, XYZ start, XYZ end, Element wall, double wallThickness)
        {
            List<XYZ> intersections = new List<XYZ>();
            XYZ direction = (end - start).Normalize();
            double segmentLength = start.DistanceTo(end);

            // Находим все пересечения вдоль сегмента
            IList<ReferenceWithContext> references = intersector.Find(start, direction);

            foreach (ReferenceWithContext refWithContext in references)
            {
                Reference reference = refWithContext.GetReference();
                XYZ intersectionPoint = reference.GlobalPoint;
                double distanceFromStart = start.DistanceTo(intersectionPoint);

                if (distanceFromStart > segmentLength)
                    continue;

                // Получаем поверхность стены
                Wall wallElement = wall as Wall;
                Face wallFace = wallElement?.GetGeometryObjectFromReference(reference) as Face;

                if (wallFace == null)
                    continue;

                // Определяем, является ли стена изогнутой
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
        /// Стена. Метод для получения размеров стены
        /// </summary>
        double GetWallThickness(Element wall)
        {
            // Пробуем получить толщину через параметр WALL_ATTR_WIDTH_PARAM
            double wallThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? 0;

            // Если не удалось, пробуем через WallType и CompoundStructure
            if (wallThickness == 0 && wall is Wall actualWall)
            {
                WallType wallType = actualWall.Document.GetElement(actualWall.GetTypeId()) as WallType;
                if (wallType != null && wallType.GetCompoundStructure() != null)
                {
                    wallThickness = wallType.GetCompoundStructure().GetWidth();
                }
            }

            return wallThickness;
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
        public void SetHoleDimensions(Transaction tx, Wall wall, Element intersectingElement, FamilyInstance holeInstance, XYZ holeLocation, double widthElement, double heightElement)
        {
            Forms.sChoiseHoleIndentation sChoiseHoleIndentation = new Forms.sChoiseHoleIndentation();

            // Открываем диалоговое окно для выбора отступа
            if (sChoiseHoleIndentation.ShowDialog() == true)
            {
                double offset = sChoiseHoleIndentation.SelectedOffset;

                Parameter widthParam = holeInstance.LookupParameter("Ширина");
                Parameter heightParam = holeInstance.LookupParameter("Высота");
                Parameter heightOtherParam = holeInstance.LookupParameter("КП_Р_Высота");
                Parameter diametrParam = holeInstance.LookupParameter("Диаметр");

                double internalHeight = UnitUtils.ConvertToInternalUnits(heightElement + offset, UnitTypeId.Millimeters);
                double internalWidth = UnitUtils.ConvertToInternalUnits(widthElement + offset, UnitTypeId.Millimeters);
                double internalDiametr = UnitUtils.ConvertToInternalUnits(Math.Max(widthElement, heightElement) + offset, UnitTypeId.Millimeters);

                if (_departmentHoleName == "АР")
                {
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
                else if (_departmentHoleName == "ИОС")
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

                    // Получаем кривую стены
                    Curve wallCurve = (wall.Location as LocationCurve).Curve;
                    XYZ wallDirection;

                    if (wallCurve is Line line) // Прямая стена
                    {
                        wallDirection = line.Direction;
                    }
                    else if (wallCurve is Arc arc) // Изогнутая стена
                    {
                        XYZ arcCenter = arc.Center;

                        XYZ radialVector = (holeLocation - arcCenter).Normalize();

                        wallDirection = new XYZ(-radialVector.Y, radialVector.X, 0);
                    }
                    else
                    {
                        return;
                    }

                    double angle = Math.Atan2(wallDirection.Y, wallDirection.X);
                    Line rotationAxis = Line.CreateBound(holeLocation, holeLocation + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(wall.Document, holeInstance.Id, rotationAxis, angle);
                    double wallThickness = GetWallThickness(wall);


                    if (wallThickness > 0)
                    {
                        XYZ wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();

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

                // Запись данных в хранилище
                string wallIdString = wall.Id.IntegerValue.ToString();
                string intersectingElementIdString = intersectingElement.Id.IntegerValue.ToString();

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
                    "10",
                    "Отверстие создано"
                );

                // Очистка панели информации перед завершением транзакции
                DockableManagerForm.Instance?.InfoHolePanel.Children.Clear();

            }
            else
            {
                transactionStatus = false;
                TaskDialog.Show("Внимание", "Операция отDменена пользователем.");
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