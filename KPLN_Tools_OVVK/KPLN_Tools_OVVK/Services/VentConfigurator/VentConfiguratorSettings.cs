using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal static class VentConfiguratorSettings
    {
        internal const string PluginName = "Конфигуратор вент. установок";
        internal const string UnknownProjectName = "Неизвестно";
        internal const string ProjectDatabasePath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";
        internal const string ProjectFamiliesRoot = @"Z:\Отдел BIM\07_Вспомогательное\Семейства\03_Вентустановки";
        internal const string ProjectFamilyFileName = "550_Универсальная установка_Одноуровневая_(Об).rfa";
        internal const string BaseFamilyTypeName = "имяСистемы(при необходимости)_имяУстановкиПоПодборке";
        // Известный демонстрационный тип исходного семейства сохраняется без изменений.
        internal const string ConfiguredFamilyTypeName = "П6_KT 2,2-250911903-02.02-K-O-P-A";
        // Обычный путь Windows: обратная косая черта перед '_' в сообщении
        // может быть экранированием Markdown, а не разделителем папок.
        internal const string SourceFamilyPath = @"X:\BIM\3_Семейства\4_ОВиК\4_ОВ2_Вентиляция\550-596_Инженерное оборудование\550_Универсальная установка_Одноуровневая_(Об).rfa";
        internal const string LiteralSourceFamilyPath = @"X:\BIM\3\_Семейства\4\_ОВиК\4\_ОВ2\_Вентиляция\550-596\_Инженерное оборудование\550\_Универсальная установка\_Одноуровневая\_(Об).rfa";

        internal static string ResolveSourcePath()
        {
            // Сначала проверяем буквальный вариант, затем вариант без Markdown-экранирования.
            return File.Exists(LiteralSourceFamilyPath) ? LiteralSourceFamilyPath : SourceFamilyPath;
        }

        internal static bool CanUpdateProjectFamily()
        {
            // То же условие, что в Module.cs: SQLiteMainService.CurrentUserDBSubDepartment.Id == 8.
            // Сервис берём из уже загруженных модулей KPLN: отдельная копия БД и новая зависимость не нужны.
            try
            {
                var services = new List<Type>();
                foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies().Where(a => a.GetName().Name.StartsWith("KPLN", StringComparison.OrdinalIgnoreCase)))
                {
                    Type[] types;
                    try
                    {
                        types = assembly.GetTypes();
                    }
                    catch (System.Reflection.ReflectionTypeLoadException ex)
                    {
                        types = ex.Types;
                    }

                    services.AddRange(types.Where(t => t != null && t.Name == "SQLiteMainService"));
                }

                if (services.Count != 1)
                    return false;
                object department = ReadAccessMember(services[0], null, "CurrentUserDBSubDepartment", true);
                if (department == null)
                    return false;
                object id = ReadAccessMember(department.GetType(), department, "Id", false);
                return id != null && Convert.ToInt32(id, CultureInfo.InvariantCulture) == 8;
            }
            catch
            {
                return false;
            } // Неопределённый пользователь или недоступный сервис не дают права обновления.
        }

        private static object ReadAccessMember(Type type, object instance, string name, bool isStatic)
        {
            var flags = System.Reflection.BindingFlags.Public | (isStatic ? System.Reflection.BindingFlags.Static : System.Reflection.BindingFlags.Instance);
            var property = type.GetProperty(name, flags);
            if (property != null)
                return property.GetValue(instance, null);
            var field = type.GetField(name, flags);
            return field == null ? null : field.GetValue(instance);
        }

        internal static string CleanDatabaseValue(string value)
        {
            string result = (value ?? string.Empty).Trim();
            return string.Equals(result, "Empty", StringComparison.OrdinalIgnoreCase) ? string.Empty : result;
        }

        internal static string MakeFileName(string projectCode)
        {
            string code = CleanDatabaseValue(projectCode);
            if (code.Length == 0)
                code = UnknownProjectName;
            char[] invalid = Path.GetInvalidFileNameChars();
            string safeName = new string(code.Select(c => invalid.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).TrimEnd('.', ' ');
            if (safeName.Length > 80)
                safeName = safeName.Substring(0, 80).TrimEnd('.', ' ');
            if (safeName.Length == 0)
                safeName = UnknownProjectName;
            return "550_Универсальная установка_Одноуровневая_(" + safeName + ").rfa";
        }
    }
}
