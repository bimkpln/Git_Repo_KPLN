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
            XYZ holeLocation;

            try
            {
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

                // Транзакция размещение отверстия
                using (Transaction tx = new Transaction(doc, $"KPLN. Разместить отверстие {GetFamilyFileName(_departmentHoleName, _holeTypeName)}"))
                {
                    tx.Start();

                    // Загрузка семейства
                    Family family = null; // Для избежания повтоной загрузки

                    FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Family));

                    foreach (Family existingFamily in collector)
                    {
                        if (existingFamily.Name == Path.GetFileNameWithoutExtension(familyPath))
                        {
                            family = existingFamily;
                            break;
                        }
                    }

                    // Если семейство не найдено, загружаем его
                    if (family == null)
                    {
                        if (!doc.LoadFamily(familyPath, out family))
                        {
                            tx.RollBack();
                            TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                            return;
                        }
                    }

                    // Ищем первый доступный типоразмер семейства
                    FamilySymbol holeSymbol = null; // Для чтения типоразмеров

                    foreach (ElementId id in family.GetFamilySymbolIds())
                    {
                        holeSymbol = doc.GetElement(id) as FamilySymbol;
                        break;
                    }

                    if (holeSymbol == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                        return;
                    }

                    // Проверяем, активен ли типоразмер, если нет — активируем
                    if (!holeSymbol.IsActive)
                    {
                        holeSymbol.Activate();
                        doc.Regenerate();
                    }

                    try
                    {
                        holeLocation = uiDoc.Selection.PickPoint("Выберите точку для размещения отверстия");
                    }
                    catch (Exception)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Ошибка", "Отмена выбора точки. Отверстие не создано.");
                        return;
                    }

                    FamilyInstance holeInstance = doc.Create.NewFamilyInstance(holeLocation, holeSymbol, _selectedElement, StructuralType.NonStructural);

                    tx.Commit();
                }

                // Успешное завершение
                TaskDialog.Show("Готово", "Операция успешно завершена.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла непредвиденная ошибка:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Выбор файла семейства в зависимости от отдела и типа отверстия
        /// </summary>
        /// <returns></returns>
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
    }
}