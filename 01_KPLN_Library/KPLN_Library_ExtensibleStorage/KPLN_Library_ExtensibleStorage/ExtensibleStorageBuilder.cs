using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;

namespace KPLN_Library_ExtensibleStorage
{
    /// <summary>
    /// Класс по работе с Extensible Storage
    /// </summary>
    public class ExtensibleStorageBuilder
    {
        /// <summary>
        /// Основной разделитель между данными, которые пишуться в ExtensibleStorage
        /// </summary>
        private readonly string _mainDataSeparator = "~ZhvBlr~";

        public ExtensibleStorageBuilder(Guid guid, string fieldName, string storageName)
        {
            Guid = guid;
            FieldName = fieldName;
            StorageName = storageName;
        }

        public ExtensibleStorageBuilder(Guid guid, string[] fieldNames, string storageName)
        {
            Guid = guid;
            FieldNames = fieldNames;
            StorageName = storageName;
        }

        /// <summary>
        /// Guid ExtensibleStorage
        /// </summary>
        public Guid Guid { get; private set; }

        /// <summary>
        /// Имя поля в ExtensibleStorage
        /// </summary>
        public string FieldName { get; private set; }

        /// <summary>
        /// Имя полей в ExtensibleStorage
        /// </summary>
        public string[] FieldNames { get; private set; }

        /// <summary>
        /// Имя ExtensibleStorage
        /// </summary>
        public string StorageName { get; private set; }

        /// <summary>
        /// Возвращает спец. класс ResultMessage из ExtensibleStorage для элемента
        /// </summary>
        /// <param name="elem">Элемент, из которого получаем ResultMessage</param>
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
                                Description = CleanDataByMainSeparator(entity.Get<string>(FieldName)),
                            };
                    }
                }
            }

            return new ResultMessage()
            {
                CurrentStatus = MessageStatus.Error,
                Description = $"Данные отсутствуют (не запускался)",
            };
        }

        /// <summary>
        /// Установить данные: последний запуск
        /// </summary>
        /// <param name="elem">Элемент, в который будет записы данные</param>
        /// <param name="userName">Имя пользователя</param>
        /// <param name="dateTime">Дата</param>
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

        /// <summary>
        /// Установить данные: текстовая пометка с разделителем в формате "Id-элемента+Разделитель+Текстовое описание"
        /// </summary>
        /// <param name="elem">Элемент, в который будет записы данные, и из которого получаю Id</param>
        /// <param name="descr">Текстовое описание</param>
        public void SetStorageData_TextData(string fieldName, Element elem, string descr)
        {
            Schema sch;
            Entity entity = CheckStorageExists(elem);
            if (entity == null)
            {
                sch = CreateSchema();
                entity = new Entity(sch);
            }

            entity.Set<string>(fieldName, $"{descr}");
            elem.SetEntity(entity);
        }

        /// <summary>
        /// Установить данные: текстовая пометка с разделителем в формате "Id-элемента+Разделитель+Текстовое описание"
        /// </summary>
        /// <param name="elem">Элемент, в который будет записы данные, и из которого получаю Id</param>
        /// <param name="userName">Имя пользователя</param>
        /// <param name="descr">Текстовое описание</param>
        public void SetStorageDataWithSeparator_TextLog(Element elem, string userName, string descr)
        {
            Schema sch;
            Entity entity = CheckStorageExists(elem);
            if (entity == null)
            {
                sch = CreateSchema();
                entity = new Entity(sch);
            }

            entity.Set<string>(FieldName, $"{elem.Id}{_mainDataSeparator}{userName}: {descr}");
            elem.SetEntity(entity);
        }

        /// <summary>
        /// Очистить данные в ExtensibleStorage
        /// </summary>
        /// <param name="elem">Элемент, который будет очищен (string.Empty)</param>
        public void DropStorageData_TextLog(Element elem)
        {
            Entity entity = CheckStorageExists(elem);
            if (entity != null)
            {
                entity.Set<string>(FieldName, string.Empty);
                elem.SetEntity(entity);
            }
        }

        /// <summary>
        /// Проверить, совпадают ли данные в ExtensibleStorage с ключевой пометкой (например: Id, служебный комментарий и т.п.)
        /// </summary>
        /// <param name="elem">Элемент, который нужно проверить</param>
        /// <param name="checkData">Данные, которые проверяем (ключевая пометка)</param>
        public bool CheckStorageDataContains_TextLog(Element elem, string checkData)
        {
            Schema sch = Schema.Lookup(Guid);
            if (sch == null) return false;

            Entity entity = CheckStorageExists(elem);
            if (entity == null) return false;

            string dataStream = entity.Get<string>(FieldName);
            if (dataStream == null) return false;

            if (dataStream.Contains(checkData)) return true;

            return false;
        }

        /// <summary>
        /// Проверяет, есть ли хоть какие-то текстовые данные в ExtensibleStorage
        /// </summary>
        /// <param name="elem">Элемент проверки</param>
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

        /// <summary>Прочитать строку без обработки разделителей. Нет записи — null.</summary>
        public string GetStorageData_TextData(string fieldName, Element elem)
        {
            Entity entity = CheckStorageExists(elem);
            if (entity == null)
                return null;
            Field field = entity.Schema.GetField(fieldName);
            if (field == null || field.ValueType != typeof(string))
                throw new InvalidOperationException("Поле Extensible Storage отсутствует или не является строкой: " + fieldName);
            return entity.Get<string>(field);
        }

        /// <summary>Прочитать несколько строк за одно обращение к сущности. Нет записи — значения null.</summary>
        public IDictionary<string, string> GetStorageData_TextData(Element elem, IEnumerable<string> fieldNames)
        {
            if (fieldNames == null)
                throw new ArgumentNullException(nameof(fieldNames));

            Entity entity = CheckStorageExists(elem);
            var values = new Dictionary<string, string>();

            foreach (string fieldName in fieldNames)
            {
                if (entity == null)
                {
                    values[fieldName] = null;
                    continue;
                }

                Field field = entity.Schema.GetField(fieldName);
                if (field == null || field.ValueType != typeof(string))
                    throw new InvalidOperationException("Поле Extensible Storage отсутствует или не является строкой: " + fieldName);

                values[fieldName] = entity.Get<string>(field);
            }

            return values;
        }

        /// <summary>Записать несколько строк в одну сущность в открытой транзакции вызывающего кода.</summary>
        public void SetStorageData_TextData(Element elem, IDictionary<string, string> values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            Entity entity = CheckStorageExists(elem) ?? new Entity(CreateSchema());
            foreach (KeyValuePair<string, string> value in values)
            {
                Field field = entity.Schema.GetField(value.Key);
                if (field == null || field.ValueType != typeof(string))
                    throw new InvalidOperationException("Поле Extensible Storage отсутствует или не является строкой: " + value.Key);
                entity.Set<string>(field, value.Value ?? string.Empty);
            }
            elem.SetEntity(entity);
        }

        /// <summary>Получить общую схему или создать её при первой записи.</summary>
        private Schema CreateSchema()
        {
            // Схема общая для процесса Revit, даже когда на этом элементе ещё нет сущности.
            Schema existing = Schema.Lookup(Guid);
            if (existing != null)
                return existing;
            SchemaBuilder builder = new SchemaBuilder(Guid);
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);

            if (FieldNames != null && FieldNames.Length > 0)
            {
                foreach (string fieldName in FieldNames)
                {
                    builder.AddSimpleField(fieldName, typeof(string));
                }
            }
            else
                builder.AddSimpleField(FieldName, typeof(string));

            builder.SetSchemaName(StorageName);

            Schema sch = builder.Finish();
            return sch;
        }

        /// <summary>
        /// Проверка наличия ExtensibleStorage для элемента
        /// </summary>
        /// <param name="elem">Элемент проверки</param>
        private Entity CheckStorageExists(Element elem)
        {
            Schema sch = Schema.Lookup(Guid);
            if (sch == null)
                return null;

            Entity ent = elem.GetEntity(sch);
            if (ent.IsValid())
                return ent;

            return null;
        }

        /// <summary>
        /// Очистка данных в ExtensibleStorage от ключевой пометки (например: Id, служебный комментарий и т.п.), если она есть
        /// </summary>
        /// <param name="data">Данные для очистки</param>
        private string CleanDataByMainSeparator(string data)
        {
            string[] splitArray = data.Split(new string[] {_mainDataSeparator}, StringSplitOptions.None);
            if (splitArray.Length > 1) return splitArray[1];
            else return splitArray[0];
        }
    }
}
