﻿using KPLN_Library_Forms.UI;
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
                #region Предобработка переданного пути
                pathFrom = RemoveSubstringIfExists(pathFrom, "http:");

                string[] pathParts = pathFrom.Split('\\');
                if (pathParts.Length < 2)
                    pathParts = pathFrom.Split('/');

                if (pathParts.Length < 2)
                    return null;
                
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
                #endregion

                try
                {
                    RevitServer server = new RevitServer(rsHostName, int.Parse(revitVersion));
                    // Проверяю ссылку на конечный файл. Добавляю файл
                    if (pathFrom.ToLower().Contains("rvt"))
                    {
                        string folderPath = string.Join("\\", pathParts, 3, pathPartsLenght - 4);
                        FolderContents folderContents = server.GetFolderContents(folderPath);
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

        private static string RemoveSubstringIfExists(string originalString, string substring)
        {
            int index = originalString.IndexOf(substring);
            if (index != -1)
            {
                // Удаляем подстроку, если она найдена
                return originalString.Remove(index, substring.Length);
            }

            // Возвращаем исходную строку, если подстрока не найдена
            return originalString;
        }
    }
}
