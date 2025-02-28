using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_HoleManager.Commands
{
    public class _iDataProcessor
    {
        // Список всех имён экземпляров семейств отверстия
        public static List<string> familyInstanceNameList = new List<string>
        {
            "199_Отверстие прямоугольное_(Об_Стена)", "199_Отверстие круглое_(Об_Стена)",
            "501_ЗИ_Отверстие_Прямоугольное_Стена_(Об)", "501_ЗИ_Отверстие_Круглое_Стена_(Об)"
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
                    familyInstanceNameList.Contains(instance.Symbol.Family.Name)) 
                .Select(fi => fi.Id)
                .ToList();

            return instanceIds;
        }

        // Получение статуса всех заданий на отверстия
        public static List<int> StatusHoleTask(Document doc, List<ElementId> familyInstanceIds)
        {
            // Создаем список, чтобы хранить статистику по статусам: "Без статуса", "Утверждено", "Предупреждения", "Ошибки"
            List<int> statusCounts = new List<int> { 0, 0, 0, 0 };

            foreach (var id in familyInstanceIds)
            {
                // Получаем экземпляр семейства
                FamilyInstance instance = doc.GetElement(id) as FamilyInstance; ;
                if (instance == null)
                    continue;

                // Получаем список сообщений
                List<string> messages = ExtensibleStorageHelper.GetChatMessages(instance);

                if (messages.Count > 0)
                {
                    // Берем последнее сообщение
                    string lastMessage = messages.Last();

                    // Разделяем сообщение на части по разделителю
                    string[] messageParts = lastMessage.Split(new string[] { Commands.ExtensibleStorageHelper.Separator }, StringSplitOptions.None);

                    string status = messageParts[9];

                    switch (status)
                    {
                        case "Без статуса":
                            statusCounts[0]++;
                            break;
                        case "Утверждено":
                            statusCounts[1]++;
                            break;
                        case "Предупреждения":
                            statusCounts[2]++;
                            break;
                        case "Ошибки":
                            statusCounts[3]++;
                            break;
                    }
                    
                }
            }

            return statusCounts;
        }

        // Получение всех сообщений для заданий на отверстия
        public static List<List<string>> GetHoleTaskMessages(Document doc, List<ElementId> familyInstanceIds)
        {
            List<List<string>> holeTaskMessages = new List<List<string>>();

            foreach (var id in familyInstanceIds)
            {
                // Получаем экземпляр семейства
                FamilyInstance instance = doc.GetElement(id) as FamilyInstance;
                if (instance == null)
                    continue;

                // Получаем список сообщений
                List<string> messages = ExtensibleStorageHelper.GetChatMessages(instance);

                if (messages.Count > 0)
                {
                    // Берем последнее сообщение
                    string lastMessage = messages.Last();

                    // Разделяем сообщение на части по разделителю
                    string[] messageParts = lastMessage.Split(new string[] { Commands.ExtensibleStorageHelper.Separator }, StringSplitOptions.None);
                    holeTaskMessages.Add(messageParts.ToList());
                }
            }

            return holeTaskMessages;
        }

    }
}
