using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

namespace KPLN_HoleManager.Commands
{
    public class ExtensibleStorageHelper
    {
        public static readonly Guid SchemaGuid = new Guid("87654321-4321-4321-4321-210987654321");
        public const string FieldName = "HoleChatMessagesKPLN";
        public const string Separator = "||||||";

        // Создание схемы
        private static Schema GetOrCreateSchema()
        {
            Schema schema = Schema.Lookup(SchemaGuid);

            if (schema != null)
                return schema;

            SchemaBuilder schemaBuilder = new SchemaBuilder(SchemaGuid);
            schemaBuilder.SetSchemaName("HoleChatStorageKPLN");
            schemaBuilder.AddArrayField(FieldName, typeof(string));

            return schemaBuilder.Finish();
        }

        // Метод добавления информации в экземпляр семейства отверстия
        public static void AddChatMessage(FamilyInstance instance, string date, string userName, string fromDepartment, string toDepartment, string iElementIdString, string status, string message)
        {
            // Получаем или создаем схему
            Schema schema = GetOrCreateSchema();

            // Получаем сущность экземпляра семейства
            Entity entity = instance.GetEntity(schema);

            // Если сущность недействительна, создаем новую
            if (!entity.IsValid())
            {
                entity = new Entity(schema); // Создаем новую сущность с правильной схемой
            }

            // Получаем список сообщений
            IList<string> messages = entity.Get<IList<string>>(schema.GetField(FieldName)) ?? new List<string>();

            // Формируем новое сообщение
            string newMessage = string.Join(Separator, date, userName, fromDepartment, toDepartment, iElementIdString, status, message);
            messages.Add(newMessage);

            // Устанавливаем обновленный список сообщений в сущность
            entity.Set(schema.GetField(FieldName), messages);
            instance.SetEntity(entity); // Применяем изменения
        }

        // Метод получения информации из экземпляра семейства отверстия
        public static List<string> GetChatMessages(FamilyInstance instance)
        {
            Schema schema = Schema.Lookup(SchemaGuid);
            if (schema == null)
                return new List<string>();

            Entity entity = instance.GetEntity(schema);
            if (!entity.IsValid())
                return new List<string>();

            return entity.Get<IList<string>>(schema.GetField(FieldName))?.ToList() ?? new List<string>();
        }

        // Тестовая функция
        public static void showTestMessage(FamilyInstance instance)
        {
            // Получаем сообщения через существующую функцию GetChatMessages
            var chatMessages = ExtensibleStorageHelper.GetChatMessages(instance);

            // Формируем сообщение для TaskDialog
            string messageContent = $"Количество сообщений: {chatMessages.Count}\n\n";

            if (chatMessages.Count > 0)
            {
                // Если есть сообщения, выводим их
                messageContent += string.Join("\n", chatMessages.Select(msg => msg.Replace("||||||", "\n")));
            }
            else
            {
                // Если сообщений нет
                messageContent += "Нет сообщений.";
            }

            // Отображаем в диалоговом окне
            TaskDialog.Show("История изменений", messageContent);
        }
    }
}