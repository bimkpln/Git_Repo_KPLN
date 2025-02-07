using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        private readonly Element _selectedElement;
        private readonly string _departmentHoleName;
        private readonly string _sendingDepartmentHoleName;
        private readonly string _holeTypeName;

        public PlaceHoleOnWallCommand(string userFullName, string departmentName, Element selectedElement, string departmentHoleName, string sendingDepartmentHoleName, string holeTypeName)
        {
            _userFullName = userFullName;
            _departmentName = departmentName;
            _selectedElement = selectedElement;
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
                // Запрашиваем выбор элемента, который пересекает стену
                TaskDialog.Show("Выбор элемента", "Для продолжения работы выберите элемент, который пересекается со стеной.");
                Reference selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Выберите элемент, который пересекается со стеной.");
                Element intersectingElement = doc.GetElement(selectedRef);

                if (intersectingElement == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось выбрать элемент для пересечения.");
                    return;
                }

                // Определяем точку пересечения
                var (holeLocation, holeWidth, holeHeight) = GetIntersectionData(_selectedElement, intersectingElement);

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
                    FamilyInstance holeInstance = doc.Create.NewFamilyInstance(holeLocation, holeSymbol, _selectedElement, StructuralType.NonStructural);

                    if (holeInstance == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Ошибка при создании отверстия.");
                        return;
                    }

                    // Устанавливаем размеры отверстия
                    SetHoleSize(holeInstance, holeWidth, holeHeight);

                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
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

        // Метод поиска точки установки отверстия
        private (XYZ, double, double) GetIntersectionData(Element wall, Element intersectingElement)
        {
            BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);
            BoundingBoxXYZ elemBox = intersectingElement.get_BoundingBox(null);

            if (wallBox == null || elemBox == null)
            {
                return (null, 0, 0);
            }

            // Определяем границы пересечения
            double minX = Math.Max(wallBox.Min.X, elemBox.Min.X);
            double maxX = Math.Min(wallBox.Max.X, elemBox.Max.X);

            double minY = Math.Max(wallBox.Min.Y, elemBox.Min.Y);
            double maxY = Math.Min(wallBox.Max.Y, elemBox.Max.Y);

            double minZ = Math.Max(wallBox.Min.Z, elemBox.Min.Z);
            double maxZ = Math.Min(wallBox.Max.Z, elemBox.Max.Z);

            // Проверяем, есть ли реальное пересечение
            if (minX > maxX || minY > maxY || minZ > maxZ)
            {
                return (null, 0, 0);
            }

            // Центр пересечения
            XYZ intersectionCenter = new XYZ(
                (minX + maxX) / 2,
                (minY + maxY) / 2,
                (minZ + maxZ) / 2
            );

            // Размеры пересечения (ширина и высота)
            double width = maxX - minX;
            double height = maxZ - minZ;

            return (intersectionCenter, width, height);
        }

        // Метод установки размера отверстия
        private void SetHoleSize(FamilyInstance hole, double width, double height)
        {
            Parameter widthParam = hole.LookupParameter("Width");
            Parameter heightParam = hole.LookupParameter("Height");

            if (widthParam != null && widthParam.IsReadOnly == false)
            {
                widthParam.Set(width);
            }

            if (heightParam != null && heightParam.IsReadOnly == false)
            {
                heightParam.Set(height);
            }
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