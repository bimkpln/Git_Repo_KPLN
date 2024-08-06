using KPLN_Library_Forms.UI;
using KPLN_Tools.Common.LinkManager;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_Tools.Common
{
    /// <summary>
    /// Кастомный сервис по работе с RS
    /// </summary>
    internal class EnvironmentService
    {
        /// <summary>
        /// Получить коллекцию файлов по указанному пути, в том числе и для Ревит-Сервер
        /// </summary>
        public static List<string> GetFilePathesFromPath(string pathFrom, string revitVersion)
        {
            List<string> fileFromPathes = new List<string>();

            // Проверяю, что это файл, если нет - то нужно забрать ВСЕ файлы из папки
            if (System.IO.File.Exists(pathFrom))
                fileFromPathes.Add(pathFrom);
            // Проверяю, что это папка, если нет - то нужно забрать ВСЕ файлы из ревит-сервера
            else if (Directory.Exists(pathFrom))
                fileFromPathes = Directory.GetFiles(pathFrom, "*" + ".rvt").ToList<string>();
            // Обработка Revit-Server, чтобы забрать файл или ВСЕ файлы из папки
            // https://www.nuget.org/packages/RevitServerAPILib
            else
            {

                string[] pathParts = pathFrom.Split('\\');

                string rsHostName = pathParts[2];
                int pathPartsLenght = pathParts.Length;
                if (rsHostName == null)
                {
                    CustomMessageBox cmb = new CustomMessageBox(
                        "Ошибка",
                        $"Ошибка заполнения пути для копирования с Revit-Server: ({pathFrom}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'");
                    cmb.ShowDialog();
                    return null;
                }

                try
                {
                    RevitServer server = new RevitServer(rsHostName, int.Parse(revitVersion));
                    // Проверяю ссылку на конечный файл. Добавляю файл
                    if (pathFrom.ToLower().Contains("rvt"))
                    {
                        FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 4));
                        foreach (var model in folderContents.Models)
                        {
                            if (model.Name == pathParts[pathPartsLenght - 1])
                            {
                                fileFromPathes.Add($"RSN:{pathFrom}");
                                break;
                            }
                        }
                    }
                    // Значит ссылка на папку. Добавляю файлы
                    else
                    {
                        FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 3));
                        foreach (var model in folderContents.Models)
                        {
                            fileFromPathes.Add($"RSN:{pathFrom}{model.Name}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    CustomMessageBox cmb = new CustomMessageBox(
                        "Ошибка",
                        $"Ошибка открытия Revit-Server ({pathFrom}):\n{ex.Message}");
                    cmb.ShowDialog();
                    return null;
                }
            }

            return fileFromPathes;
        }

        /// <summary>
        /// Подготовить коллекцию LinkChangeEntity из предоставленных путей
        /// </summary>
        public static List<LinkManagerEntity> PrepareLCEntityByPathes(List<string> fileFromPathes)
        {
            List<LinkManagerEntity> result = new List<LinkManagerEntity>();
            
            foreach (string path in fileFromPathes)
            {
                string[] pathParts = path.Split('\\');
                string modelName = pathParts.Where(x => x.EndsWith("rvt")).FirstOrDefault();
                if (modelName == null)
                    throw new Exception($"Не удалось получить имя файла по пути: {path}");
                else
                    result.Add(new LinkManagerEntity(modelName, path));
            }

            return result;
        }
    }
}
