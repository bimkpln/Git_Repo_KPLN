using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KPLN_Scoper.Services
{
    /// <summary>
    /// Сервис перезаписи данных в ini-файлах ревит
    /// </summary>
    public class INIFileService
    {
        private const int SIZE = 1024; //Максимальный размер (для чтения значения из файла)
        private readonly string _path = string.Empty; //Для хранения пути к INI-файлу
        private readonly string _revitVersion = string.Empty; //Версия ревит
        private readonly SQLUserInfo _dbUserInfo; //Версия ревит

        public INIFileService(SQLUserInfo dbUserInfo, string revitVersion)
        {
            _dbUserInfo = dbUserInfo;
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
            {
                throw new NullReferenceException($"По ссылке отсутсвует ini-файл: {ini_fi.FullName}");
            }

            // Запись настроек выборки
            WritePrivateString("Selection", "AllowPressAndDrag", "0", ini_fi.FullName);
            WritePrivateString("Selection", "AllowFaceSelection", "0", ini_fi.FullName);
            WritePrivateString("Selection", "AllowUnderlaySelection", "1", ini_fi.FullName);

            // Запись шаблонов ревит КПЛН
            if (_revitVersion == "2020")
            {
                WritePrivateString("DirectoriesRUS", "DefaultTemplate", GetTemplates());
                WritePrivateString("DirectoriesENU", "DefaultTemplate", GetTemplates());
            }
            else
                return false;

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

        //Пишет значение в INI-файл (по указанным секции и ключу) 
        public bool WritePrivateString(string aSection, string aKey, string aValue)
        {
            //Записать значение в INI-файл
            WritePrivateString(aSection, aKey, aValue, _path);
            return true;
        }

        //Импорт функции GetPrivateProfileString (для чтения значений) из библиотеки kernel32.dll
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateString(string section, string key, string def, StringBuilder buffer, int size, string path);

        //Импорт функции WritePrivateProfileString (для записи значений) из библиотеки kernel32.dll
        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileString")]
        private static extern int WritePrivateString(string section, string key, string str, string path);

        /// <summary>
        /// Получить коллекцию полных путей шаблонов КПЛН
        /// </summary>
        /// <returns></returns>
        private string GetTemplates()
        {
            List<string> parts = new List<string>();

            DirectoryInfo templateFolder = new DirectoryInfo(@"X:\BIM\2_Шаблоны");
            foreach (DirectoryInfo folder in templateFolder.GetDirectories())
            {
                if (_dbUserInfo.Department.Id == 1 && (folder.Name != "1_АР" && folder.Name != "0_Общие шаблоны")) { continue; }
                if (_dbUserInfo.Department.Id == 2 && (folder.Name != "2_КР" && folder.Name != "0_Общие шаблоны")) { continue; }
                if (_dbUserInfo.Department.Id == 3 && (folder.Name == "1_АР" || folder.Name == "2_КР" || folder.Name == "0_Общие шаблоны")) { continue; }
                foreach (FileInfo file in folder.GetFiles())
                {
                    if (IsCopy(file.Name) || file.Extension.ToLower() != ".rte") { continue; }
                    parts.Add(string.Format("{0} - {1}={2}", folder.Name.Split('_').Last(), OnlyName(file.Name), file.FullName));
                }
            }
            return string.Join(",", parts);
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
