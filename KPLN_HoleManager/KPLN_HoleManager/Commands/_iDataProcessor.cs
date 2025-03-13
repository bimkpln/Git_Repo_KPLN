using Autodesk.Revit.DB;
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
        public static List<int> StatusHoleTask(Document doc, List<ElementId> familyInstanceIds, string userName, string userDepartament)
        {
            // Создаем список, чтобы хранить статистику по статусам: "Без статуса", "Утверждено", "Подтверждение", "Ошибки"
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

                    string status = messageParts[10];
                
                    switch (status)
                    {
                        case "Без статуса":
                            if (messageParts[1] == userName || messageParts[3] == userDepartament || messageParts[1] == "Функция плагина" || userDepartament == "BIM") statusCounts[0]++;
                            break;
                        case "Утверждено":
                            if (messageParts[1] == userName || messageParts[3] == userDepartament || messageParts[4] == userDepartament || userDepartament == "BIM") statusCounts[1]++;
                            break;
                        case "Подтверждение":
                            if (messageParts[1] == userName || messageParts[3] == userDepartament || messageParts[4] == userDepartament || userDepartament == "BIM") statusCounts[2]++;
                            break;
                        case "Ошибки":
                            if (messageParts[1] == userName || messageParts[3] == userDepartament || messageParts[4] == userDepartament || userDepartament == "BIM") statusCounts[3]++;
                            break;
                    }
                                    
                }
            }

            return statusCounts;
        }

        // Получение всех сообщений для заданий на отверстия из всех элементов (получаем последнее)
        public static List<List<string>> GetHoleLastTaskMessages(Document doc, List<ElementId> familyInstanceIds)
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

        // Получение всех сообщений для заданий на отверстия из одного элемента (получаем всё)
        public static List<List<string>> GetHoleTaskMessages(Document doc, string holeID)
        {
            List<List<string>> holeTaskMessages = new List<List<string>>();

            if (int.TryParse(holeID, out int elementIdValue))
            {
                ElementId elementId = new ElementId(elementIdValue);
                Element element = doc.GetElement(elementId);

                if (element is FamilyInstance instance)
                {
                    List<string> messages = ExtensibleStorageHelper.GetChatMessages(instance);

                    foreach (string message in messages)
                    {
                        string[] messageParts = message.Split(new string[] { Commands.ExtensibleStorageHelper.Separator }, StringSplitOptions.None);

                        holeTaskMessages.Add(messageParts.ToList());
                    }
                }
            }

            return holeTaskMessages;
        }



    }
}
