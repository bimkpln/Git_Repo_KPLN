using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace KPLN_HoleManager.ExternalCommand
{
    public class PlaceHoleOnWallCommand
    {
        private readonly Element _selectedElement;
        private readonly string _departmentHoleName;
        private readonly string _holeTypeName;

        public PlaceHoleOnWallCommand(Element selectedElement, string departmentHoleName, string holeTypeName)
        {
            _selectedElement = selectedElement;
            _departmentHoleName = departmentHoleName;
            _holeTypeName = holeTypeName;
        }

        public static void Execute(UIApplication uiApp, Element selectedElement, string departmentHoleName, string holeTypeName)
        {
            if (uiApp == null)
            {
                TaskDialog.Show("Ошибка", "Revit API не доступен.");
                return;
            }

            // Создаём экземпляр команды и передаём её в ExternalEvent
            var command = new PlaceHoleOnWallCommand(selectedElement, departmentHoleName, holeTypeName);
            _ExternalEventHandler.Instance.Raise((app) => command.Run(app));
        }

        public void Run(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Проверка параметров
                if (_selectedElement == null)
                {
                    TaskDialog.Show("Ошибка", "Не выбрана стена.");
                    return;
                }
                if (string.IsNullOrEmpty(_departmentHoleName) || string.IsNullOrEmpty(_holeTypeName))
                {
                    TaskDialog.Show("Ошибка", "Не выбраны параметры отверстия.");
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

                // Загрузка семейства
                FamilySymbol holeSymbol = LoadFamilySymbol(doc, familyPath);
                if (holeSymbol == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                    return;
                }

                // Выбор точки
                XYZ holeLocation;
                try
                {
                    holeLocation = uiDoc.Selection.PickPoint("Выберите точку для размещения отверстия");
                }
                catch (Exception)
                {
                    TaskDialog.Show("Ошибка", "Отмена выбора точки. Отверстие не создано.");
                    return;
                }

                // Размещение отверстия
                using (Transaction tx = new Transaction(doc, "Разместить отверстие"))
                {
                    tx.Start();

                    // Проверяем, активен ли типоразмер, если нет — активируем
                    if (!holeSymbol.IsActive)
                    {
                        holeSymbol.Activate();
                        doc.Regenerate();
                    }

                    holeLocation = uiDoc.Selection.PickPoint("Выберите точку для размещения отверстия");
                    FamilyInstance holeInstance = doc.Create.NewFamilyInstance(holeLocation, holeSymbol, _selectedElement, StructuralType.NonStructural);

                    tx.Commit();
                }

                // Успешное завершение
                TaskDialog.Show("Готово", "Отверстие успешно добавлено. Вы можете передвинуть его вручную.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
            }
        }

        // Выбор файла семейства
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

        // Загрузка семейства
        private FamilySymbol LoadFamilySymbol(Document doc, string familyPath)
        {
            Family family;
            using (Transaction tx = new Transaction(doc, "Загрузка семейства"))
            {
                tx.Start();
                if (!doc.LoadFamily(familyPath, out family))
                {
                    tx.RollBack();
                    return null;
                }
                tx.Commit();
            }

            if (family != null)
            {
                foreach (ElementId id in family.GetFamilySymbolIds())
                {
                    return doc.GetElement(id) as FamilySymbol;
                }
            }
            return null;
        }
    }
}