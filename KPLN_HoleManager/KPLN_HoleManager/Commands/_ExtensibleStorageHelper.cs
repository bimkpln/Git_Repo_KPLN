using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

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

        /// <summary>
        /// Метод добавления информации в экземпляр семейства отверстия
        /// </summary>
        public static void AddChatMessage(FamilyInstance holeInstance, string date, string userName, string department, string fromDepartment, string toDepartment, string wallIdString, string iElementIdString, string status, string statusIO, string message)
        {
            // Получаем или создаем схему
            Schema schema = GetOrCreateSchema();

            // Получаем сущность экземпляра семейства
            Entity entity = holeInstance.GetEntity(schema);

            // Если сущность недействительна, создаем новую
            if (!entity.IsValid())
            {
                entity = new Entity(schema); // Создаем новую сущность с правильной схемой
            }

            string holeInstanceId = holeInstance.Id.IntegerValue.ToString();
            string holeInstanceName = $"{holeInstance.Symbol.Family.Name} - {holeInstance.Symbol.Name}";
            Transform transformHoleInstance = holeInstance.GetTransform();
            XYZ originHoleInstance = transformHoleInstance.Origin;
            string coordinatesHoleInstance = $"{originHoleInstance.X:F3},{originHoleInstance.Y:F3},{originHoleInstance.Z:F3}";

            // Получаем список сообщений
            IList<string> messages = entity.Get<IList<string>>(schema.GetField(FieldName)) ?? new List<string>();

            // Формируем новое сообщение
            string newMessage = string.Join(Separator, date, userName, department, fromDepartment, toDepartment, 
                holeInstanceId, holeInstanceName, wallIdString, iElementIdString, coordinatesHoleInstance, status, statusIO, message);
            messages.Add(newMessage);

            // Устанавливаем обновленный список сообщений в сущность
            entity.Set(schema.GetField(FieldName), messages);
            holeInstance.SetEntity(entity); // Применяем изменения
        }

        /// <summary>
        /// Метод получения информации из экземпляра семейства отверстия
        /// </summary>
        public static List<string> GetChatMessages(FamilyInstance holeInstance)
        {
            Schema schema = Schema.Lookup(SchemaGuid);
            if (schema == null)
                return new List<string>();

            Entity entity = holeInstance.GetEntity(schema);
            if (!entity.IsValid())
                return new List<string>();

            return entity.Get<IList<string>>(schema.GetField(FieldName))?.ToList() ?? new List<string>();
        }      
    }
}