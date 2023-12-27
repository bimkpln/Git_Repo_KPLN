using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KPLN_Looker.Services
{
    /// <summary>
    /// Сервис перезаписи данных в ini-файлах ревит
    /// </summary>
    public class INIFileService
    {
        private const int SIZE = 1024; //Максимальный размер (для чтения значения из файла)
        private readonly string _path;
        private readonly string _revitVersion;
        private readonly DBUser _user;

        public INIFileService(DBUser user, string revitVersion)
        {
            _user = user;
            _revitVersion = revitVersion;
            _path = string.Format(@"AppData\Roaming\Autodesk\Revit\Autodesk Revit {0}\Revit.ini", _revitVersion);
        }

        /// <summary>
        /// Запуск метода перезаписи ini-файла
        /// </summary>
        public bool OverwriteINIFile()
        {
            FileInfo ini_fi = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), _path));
            if (!ini_fi.Exists)
                throw new NullReferenceException($"По ссылке отсутсвует ini-файл: {ini_fi.FullName}");

            // Запись настроек выборки
            WritePrivateString("Selection", "AllowPressAndDrag", "0", ini_fi.FullName);
            WritePrivateString("Selection", "AllowFaceSelection", "0", ini_fi.FullName);
            WritePrivateString("Selection", "AllowUnderlaySelection", "1", ini_fi.FullName);

            // Запись шаблонов ревит КПЛН
            string templates = GetTemplates();
            if (templates != null)
            {
                WritePrivateString("DirectoriesRUS", "DefaultTemplate", templates, ini_fi.FullName);
                WritePrivateString("DirectoriesENU", "DefaultTemplate", templates, ini_fi.FullName);
            }
            else
                throw new NullReferenceException($"Ошибка поиска шаблонов для раздела: {ini_fi.FullName}");

            // Отключение журнала перемотки
            WritePrivateString("AutoCam", "SaveRewindThumbnails", "0", ini_fi.FullName);

            return true;
        }

        //Возвращает значение из INI-файла (по указанным секции и ключу) 
        public string GetPrivateString(string aSection, string aKey)
        {
            //Для получения значения
            StringBuilder buffer = new StringBuilder(SIZE);

            //Получить значение в buffer
            GetPrivateString(aSection, aKey, null, buffer, SIZE, _path);

            //Вернуть полученное значение
            return buffer.ToString();
        }

        /// <summary>
        /// Импорт функции GetPrivateProfileString (для чтения значений) из библиотеки kernel32.dll
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="def"></param>
        /// <param name="buffer"></param>
        /// <param name="size"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateString(string section, string key, string def, StringBuilder buffer, int size, string path);

        /// <summary>
        /// Импорт функции WritePrivateProfileString (для записи значений) из библиотеки kernel32.dll
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="str"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileString")]
        private static extern int WritePrivateString(string section, string key, string str, string path);

        /// <summary>
        /// Получить коллекцию полных путей шаблонов КПЛН
        /// </summary>
        /// <returns></returns>
        private string GetTemplates()
        {
            DirectoryInfo templateFolder = new DirectoryInfo(@"X:\BIM\2_Шаблоны");
            Dictionary<int, string[]> departmentKeywords = new Dictionary<int, string[]>
            {
                { 2, new string[] { "1_АР" } },
                { 3, new string[] { "2_КР" } },
                { 4, new string[] { "3_ОВиК", "0_Общие шаблоны" } },
                { 5, new string[] { "4_ВК", "0_Общие шаблоны" } },
                { 6, new string[] { "6_ЭОМ", "0_Общие шаблоны" } },
                { 7, new string[] { "5_СС", "0_Общие шаблоны" } },
                { 8, null }
            };

            List<string> tempFormat = templateFolder
                .GetDirectories()
                .Where(folder =>
                    departmentKeywords.TryGetValue(_user.SubDepartmentId, out var keywords) && (keywords == null || keywords.Any(keyword => folder.Name.Contains(keyword))))
                .SelectMany(folder => CreateTemplateFormat(folder))
                .ToList();

            return tempFormat.Any() ? string.Join(", ", tempFormat) : null;
        }

        private List<string> CreateTemplateFormat(DirectoryInfo folder)
        {
            List<string> parts = new List<string>();

            foreach (FileInfo file in folder.GetFiles())
            {
                if (!IsCopy(file.Name) && file.Extension.ToLower() == ".rte")
                    parts.Add($"{folder.Name.Split('_').Last()} - {OnlyName(file.Name)}={file.FullName}");
            }

            return parts;
        }

        private bool IsCopy(string name)
        {
            List<string> parts = name.Split('.').ToList();
            if (parts.Count <= 2) { return false; }
            parts.RemoveAt(parts.Count - 1);
            if (parts.Last().Length == 4 && parts.Last().StartsWith("0")) { return true; }
            return false;
        }

        private string OnlyName(string name)
        {
            List<string> parts = name.Split('.').ToList();
            parts.RemoveAt(parts.Count - 1);
            return string.Join(".", parts);
        }
    }
}
