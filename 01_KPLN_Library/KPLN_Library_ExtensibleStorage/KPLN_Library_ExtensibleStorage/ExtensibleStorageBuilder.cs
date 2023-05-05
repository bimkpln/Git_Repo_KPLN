using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;

namespace KPLN_Library_ExtensibleStorage
{
    /// <summary>
    /// Класс по работе с Extensible Storage
    /// </summary>
    public class ExtensibleStorageBuilder
    {
        public ExtensibleStorageBuilder(Guid guid, string name)
        {
            Guid = guid;
            FieldName = name;
            StorageName = "storage";
        }

        public ExtensibleStorageBuilder(Guid guid, string name, string storageName) : this(guid, name)
        {
            StorageName = storageName;
        }

        public Guid Guid { get; private set; }

        public string FieldName { get; private set; }

        public string StorageName { get; private set; }

        public void SetStorageData_TimeRunLog(Element elem, string userName, DateTime dateTime)
        {
            Schema sch;
            Entity entity = CheckStorageExists(elem);
            if (entity == null)
            {
                sch = CreateSchema();
                entity = new Entity(sch);
            }

            entity.Set<string>(FieldName, $"{userName}: {dateTime}");
            elem.SetEntity(entity);
        }

        public void SetStorageData_TextLog(Element elem, string userName, string descr)
        {
            Schema sch;
            Entity entity = CheckStorageExists(elem);
            if (entity == null)
            {
                sch = CreateSchema();
                entity = new Entity(sch);
            }

            entity.Set<string>(FieldName, $"{userName}: {descr}");
            elem.SetEntity(entity);
        }

        public ResultMessage GetResMessage_ProjectInfo(Document doc)
        {
            ProjectInfo pi = doc.ProjectInformation;
            Entity entity = CheckStorageExists((Element)pi);
            if (entity != null)
            {
                ResultMessage message = new ResultMessage()
                {
                    CurrentStatus = MessageStatus.Ok,
                    Description = $"{entity.Get<string>(FieldName)}",
                };
                return message;
            }
            else
            {
                ResultMessage message = new ResultMessage()
                {
                    CurrentStatus = MessageStatus.Error,
                    Description = $"Данные отсутсвуют (не запускался)",
                };
                return message;
            }
        }

        public ResultMessage GetResMessage_Element(Element elem)
        {
            Schema sch = Schema.Lookup(Guid);
            if (sch != null)
            {
                Entity entity = elem.GetEntity(sch);
                if (entity.Schema != null)
                {
                    Field field = entity.Schema.GetField(FieldName);
                    if (field != null)
                    {
                        string data = entity.Get<string>(field);
                        if (data != null && data != string.Empty)
                            return new ResultMessage()
                            {
                                CurrentStatus = MessageStatus.Ok,
                                Description = $"{entity.Get<string>(FieldName)}",
                            };
                    }
                }
            }

            return new ResultMessage()
            {
                CurrentStatus = MessageStatus.Error,
                Description = $"Данные отсутсвуют (не запускался)",
            };
        }

        /// <summary>
        /// Проверяет, есть ли хоть какие-то данные в ExtStr
        /// </summary>
        public bool IsDataExists_Text(Element elem)
        {
            Schema sch = Schema.Lookup(Guid);
            if (sch == null)
                return false;

            Entity ent = elem.GetEntity(sch);
            if (ent.Schema != null)
            {
                Field field = ent.Schema.GetField(FieldName);
                if (field != null)
                {
                    string data = ent.Get<string>(field);
                    if (data != null && data != string.Empty)
                        return true;
                }
            }

            return false;
        }

        private Schema CreateSchema()
        {
            SchemaBuilder builder = new SchemaBuilder(Guid);
            builder.SetReadAccessLevel(AccessLevel.Public);

            FieldBuilder fbName = builder.AddSimpleField(FieldName, typeof(string));
            builder.SetSchemaName(StorageName);

            Schema sch = builder.Finish();
            return sch;
        }

        private Entity CheckStorageExists(Element elem)
        {
            Schema sch = Schema.Lookup(Guid);
            if (sch == null)
                return null;

            Entity ent = elem.GetEntity(sch);
            if (ent.Schema != null)
                return ent;

            return null;
        }
    }
}
