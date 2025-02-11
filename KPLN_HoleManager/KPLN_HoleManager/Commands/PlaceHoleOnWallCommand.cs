using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace KPLN_HoleManager.ExternalCommand
{
    public class PlaceHoleOnWallCommand
    {
        private readonly string _userFullName;
        private readonly string _departmentName;

        private readonly Element _selectedWall;
        private string wallType;

        private readonly string _departmentHoleName;
        private readonly string _sendingDepartmentHoleName;
        private readonly string _holeTypeName;

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
                // Определение типа стены
                Wall wall = _selectedWall as Wall;
                LocationCurve locationCurve = wall.Location as LocationCurve;

                // Определяем тип стены
                if (locationCurve != null)
                {
                    Curve curve = locationCurve.Curve;
                    if (curve is Line)
                    {
                        wallType = "default";
                    }
                    else if (curve is Arc)
                    {
                        wallType = "ArcWall";
                    }
                    else if (wall.CurtainGrid != null)
                    {
                        wallType = "CurtainGridWall"; 
                    }
                    else if (wall.WallType.Kind == WallKind.Stacked)
                    {
                        wallType = "StackedCurtainGrid"; 
                    }
                    else if (_selectedWall is RevitLinkInstance linkInstance)
                    {
                        wallType = "LinkWall";
                    }
                }

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

                Element intersectingElement = null;

                // Выбор элемента, который пересекает стену
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

                // Определение размеров элемента, который пересекает стену
                (double width, double height, double length) = GetElementSize(intersectingElement);

                // Определяем точку пересечения
                XYZ holeLocation = GetIntersectionPoint(doc, _selectedWall, intersectingElement);

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

                using (Transaction tx = new Transaction(doc, $"KPLN. Разместить отверстие {familyFileName}"))
                {
                    tx.Start();

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

                    // Создаём отверстие
                    FamilyInstance holeInstance = doc.Create.NewFamilyInstance(holeLocation, holeSymbol, _selectedWall, StructuralType.NonStructural);

                    if (holeInstance == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Ошибка при создании отверстия.");
                        return;
                    }

                    Forms.sChoiseHoleIndentation sChoiseHoleIndentation = new Forms.sChoiseHoleIndentation();

                    // Устанавливаем размеры отверстия
                    if (sChoiseHoleIndentation.ShowDialog() == true)
                    {
                        double offset = sChoiseHoleIndentation.SelectedOffset;

                        Parameter widthParam = holeInstance.LookupParameter("Ширина");
                        Parameter heightParam = holeInstance.LookupParameter("Высота");

                        // Отверстие типа - прямоугольник
                        if (heightParam != null && !heightParam.IsReadOnly)
                        {
                            heightParam.Set(UnitUtils.ConvertToInternalUnits(height + offset, UnitTypeId.Millimeters));
                        }

                        if (widthParam != null && !widthParam.IsReadOnly)
                        {
                            widthParam.Set(UnitUtils.ConvertToInternalUnits(width + offset, UnitTypeId.Millimeters));
                        }

                        // Отверстие типа - круг
                        if (heightParam != null && !heightParam.IsReadOnly && widthParam != null && widthParam.IsReadOnly)
                        {
                            heightParam.Set(UnitUtils.ConvertToInternalUnits(Math.Max(height, width) + offset, UnitTypeId.Millimeters));
                        }
                    }
                    else
                    {
                        tx.RollBack();
                        TaskDialog.Show("Внимание", "Операция отменена пользователем.");
                        return;
                    }


                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
            }
        }








        // Метод для получения размеров элемента системы
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

        // Метод для получения значения параметра с переводом футов в метры
        private double GetParameterValue(Element element, BuiltInParameter paramId)
        {
            Parameter param = element.get_Parameter(paramId);
            if (param != null && param.HasValue)
            {
                // Используем Revit API для корректного преобразования в мм
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
            }
            return 0;
        }









        // Находим точку пересечения стены и входящего элемента
        private XYZ GetIntersectionPoint(Document doc, Element wall, Element cElement)
        {
            if (wallType == "default" || wallType == "ArcWall")
            {
                // Получаем геометрию трубы (учитываем изгибы)
                LocationCurve pipeLocation = cElement.Location as LocationCurve;
                if (pipeLocation == null)
                    return null;

                Curve pipeCurve = pipeLocation.Curve;

                // Разбиваем кривую на точки (Tessellate)
                List<XYZ> segmentPoints = pipeCurve.Tessellate() as List<XYZ>;

                if (segmentPoints == null || segmentPoints.Count < 2)
                    return null;

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


                // Создаём фильтр для нашей стены
                ReferenceIntersector intersector = new ReferenceIntersector(wall.Id, FindReferenceTarget.Face, view3D);

                XYZ closestIntersection = null;
                double minDistance = double.MaxValue;

                // Проходим по каждой паре точек
                for (int i = 0; i < segmentPoints.Count - 1; i++)
                {
                    XYZ start = segmentPoints[i];
                    XYZ end = segmentPoints[i + 1];
                    XYZ direction = (end - start).Normalize();

                    ReferenceWithContext referenceWithContext = intersector.FindNearest(start, direction);
                    if (referenceWithContext != null)
                    {
                        XYZ intersectionPoint = referenceWithContext.GetReference().GlobalPoint;

                        // Вычисляем расстояние до точки трубы
                        double distanceToPipe = start.DistanceTo(intersectionPoint);

                        // Если точка ближе - запоминаем её
                        if (distanceToPipe < minDistance)
                        {
                            minDistance = distanceToPipe;
                            closestIntersection = intersectionPoint;
                        }
                    }
                }

                // Получаем параметры стены
                double offsetBottom = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                double offsetTop = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0;
                double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0;
                bool isUnconnected = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)?.AsElementId() == ElementId.InvalidElementId;

                // Корректируем точку пересечения с учетом смещений и выступов
                if (closestIntersection != null)
                {
                    double adjustment = (isUnconnected ? unconnectedHeight : 0) - offsetBottom + offsetTop;
                    closestIntersection = new XYZ(closestIntersection.X, closestIntersection.Y, closestIntersection.Z + adjustment);
                }

                if (createdView)
                {
                    doc.Delete(view3D.Id);
                }

                return closestIntersection;
            }
            else
            {
                TaskDialog.Show("Внимание", $"Работа со стенами типа {wallType} не реализована.");
                return null;
            }
        }











        // Метод выбора файла семейства
        private string GetFamilyFileName(string department, string holeType)
        {
            var familyMap = new Dictionary<string, string>
            {
                { "АР_SquareHole", "199_AR_OSW.rfa" },
                { "АР_RoundHole", "199_AR_ORW.rfa" },
                { "КР_SquareHole", "199_STR_OSW.rfa" },
                { "КР_RoundHole", "199_STR_ORW.rfa" },
                { "ИОС_SquareHole", "501_MEP_TSW.rfa" },
                { "ИОС_RoundHole", "501_MEP_TRW.rfa" }
            };

            string key = $"{department}_{holeType}";
            return familyMap.ContainsKey(key) ? familyMap[key] : null;
        }

        // Метод загрузки семейства
        private Family LoadOrGetExistingFamily(Document doc, string familyPath)
        {
            Family family = new FilteredElementCollector(doc).OfClass(typeof(Family))
                .FirstOrDefault(f => f.Name == Path.GetFileNameWithoutExtension(familyPath)) as Family;

            if (family == null)
            {
                doc.LoadFamily(familyPath, out family);
            }
            return family;
        }

        // Метод загрузки типоразмера семейства
        private FamilySymbol GetFamilySymbol(Family family, Document doc)
        {
            foreach (ElementId id in family.GetFamilySymbolIds())
            {
                return doc.GetElement(id) as FamilySymbol;
            }
            return null;
        }
    }
}