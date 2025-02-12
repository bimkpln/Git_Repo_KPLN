using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_HoleManager.Commands
{
    class _iDataProcessor
    {
        // Список всех имён экземпляров семейств отверстия
        public static List<string> familyInstanceNameList = new List<string>
        {
            "199_AR_ORW", "199_AR_OSW", "199_STR_ORW", "199_STR_OSW", "501_MEP_TRW", "501_MEP_TSW"
        };

        // Получение всех ID экземпляров семейств отверстия
        public static List<ElementId> GetFamilyInstanceIds(Document doc)
        {
            // Создаем коллекцию всех экземпляров семейств
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance));

            // Фильтруем по имени семейства
            List<ElementId> instanceIds = collector
                .Where(fi =>
                    fi is FamilyInstance instance &&
                    familyInstanceNameList.Contains(instance.Symbol.Family.Name))  // Используем Symbol.Family.Name
                .Select(fi => fi.Id)
                .ToList();

            return instanceIds;
        }

        public static void ShowFamilyInstanceCount(Document doc, UIDocument uidoc, List<string> familyInstanceNameList)
        {
            List<ElementId> instanceIds = GetFamilyInstanceIds(doc);

            string message = $"Найдено отверстий: {instanceIds.Count}";

            // Выводим диалоговое окно
            TaskDialog.Show("Результат поиска", message);

            // Выбираем найденные элементы в Revit
            if (instanceIds.Count > 0)
            {
                uidoc.Selection.SetElementIds(instanceIds);
            }
        }


    }
}
