using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
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

                // Запрашиваем выбор элемента, который пересекает стену
                TaskDialog.Show("Выбор элемента", "Для продолжения работы выберите элемент, который пересекается с выбранной стеной.");
                Reference selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Выберите элемент, который пересекается со стеной.");
                Element intersectingElement = doc.GetElement(selectedRef);

                if (intersectingElement == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось выбрать элемент для пересечения.");
                    return;
                }

                // Определение размеров элемента, который пересекает стену
                (double width, double height, double length) = GetElementSize(intersectingElement);

                // Определяем точку пересечения
                XYZ holeLocation = GetIntersectionPoint(_selectedWall, intersectingElement);

                if (holeLocation == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить точку пересечения.");
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
        private XYZ GetIntersectionPoint(Element wall, Element intersectingElement)
        {
            if (wallType == "default")
            {
                BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);
                BoundingBoxXYZ intersectBox = intersectingElement.get_BoundingBox(null);

                if (wallBox == null || intersectBox == null)
                    return null;

                // Определяем трансформацию (если элемент находится в линке)
                Transform wallTransform = Transform.Identity;
                Transform intersectTransform = Transform.Identity;

                if (wall is RevitLinkInstance wallLink)
                {
                    wallTransform = wallLink.GetTotalTransform();
                }

                if (intersectingElement is RevitLinkInstance intersectLink)
                {
                    intersectTransform = intersectLink.GetTotalTransform();
                }

                // Применяем трансформацию
                XYZ wallMin = wallTransform.OfPoint(wallBox.Min);
                XYZ wallMax = wallTransform.OfPoint(wallBox.Max);
                XYZ intersectMin = intersectTransform.OfPoint(intersectBox.Min);
                XYZ intersectMax = intersectTransform.OfPoint(intersectBox.Max);

                // Определяем точку пересечения (центр общей области BoundingBox)
                double x = (Math.Max(wallMin.X, intersectMin.X) + Math.Min(wallMax.X, intersectMax.X)) / 2;
                double y = (Math.Max(wallMin.Y, intersectMin.Y) + Math.Min(wallMax.Y, intersectMax.Y)) / 2;
                double z = (Math.Max(wallMin.Z, intersectMin.Z) + Math.Min(wallMax.Z, intersectMax.Z)) / 2;

                // Проверяем, находится ли точка внутри стены
                if (x < wallMin.X || x > wallMax.X || y < wallMin.Y || y > wallMax.Y || z < wallMin.Z || z > wallMax.Z)
                    return null;

                return new XYZ(x, y, z);
            }
            else
            {
                TaskDialog.Show("Внимание", $"На данный момент времени работа со стенами типа {wallType} не реализована.");
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