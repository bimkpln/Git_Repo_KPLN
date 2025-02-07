using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace KPLN_HoleManager.Commands
{
    public class ChatStorageManager
    {
        private static readonly Guid SchemaGuid = new Guid("D4E7A1B2-8C3F-4F5A-91D8-2E6B7C8D9E0F");
        private static readonly string FieldMessages = "CommentMessages";  // Поле для сообщений

        public class TaskData
        {
            public string Date { get; set; }
            public string DepartmentFrom { get; set; }
            public string DepartmentTo { get; set; }
            public string UserName { get; set; }
            public string TaskStatus { get; set; }
            public string Comment { get; set; }

            public TaskData(string date, string deptFrom, string deptTo, string user, string status, string comment)
            {
                Date = date;
                DepartmentFrom = deptFrom;
                DepartmentTo = deptTo;
                UserName = user;
                TaskStatus = status;
                Comment = comment;
            }

            public override string ToString()
            {
                return $"{Date}|{DepartmentFrom}|{DepartmentTo}|{UserName}|{TaskStatus}|{Comment}";
            }

            public static TaskData FromString(string data)
            {
                string[] parts = data.Split('|');
                return parts.Length == 6 ? new TaskData(parts[0], parts[1], parts[2], parts[3], parts[4], parts[5]) : null;
            }
        }

        // Открытие получение схемы
        private static Schema GetOrCreateSchema()
        {
            Schema schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            SchemaBuilder schemaBuilder = new SchemaBuilder(SchemaGuid);
            schemaBuilder.SetSchemaName("KPLN_HoleManager");

            schemaBuilder.AddArrayField(FieldMessages, typeof(string));

            return schemaBuilder.Finish();
        }

        public static List<TaskData> ReadChat(FamilyInstance instance)
        {
            List<TaskData> chatMessages = new List<TaskData>();
            Schema schema = GetOrCreateSchema();
            Entity entity = instance.GetEntity(schema);

            if (!entity.IsValid()) return chatMessages;

            IList<string> messages = entity.Get<IList<string>>(schema.GetField(FieldMessages));

            foreach (string message in messages)
            {
                TaskData task = TaskData.FromString(message);
                if (task != null)
                    chatMessages.Add(task);
            }

            return chatMessages;
        }

        public static void WriteChat(Document doc, FamilyInstance instance, List<TaskData> chatMessages, Transaction transaction = null)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = new Entity(schema);

            IList<string> messages = new List<string>();

            foreach (var message in chatMessages)
            {
                messages.Add(message.ToString());
            }

            entity.Set(schema.GetField(FieldMessages), messages);

            if (transaction != null)
            {
                instance.SetEntity(entity);
            }
            else
            {
                using (Transaction trans = new Transaction(doc, "KPLN. Добавлен комментарий к заданию"))
                {
                    trans.Start();
                    instance.SetEntity(entity);
                    trans.Commit();
                }
            }
        }

        public static void AddMessage(Document doc, FamilyInstance instance, TaskData message, Transaction transaction)
        {
            List<TaskData> chat = ReadChat(instance);
            chat.Add(message);  
            WriteChat(doc, instance, chat, transaction);
        }
    }
}