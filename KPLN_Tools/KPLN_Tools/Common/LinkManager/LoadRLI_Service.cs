using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace KPLN_Tools.Common.LinkManager
{
    /// <summary>
    /// Сервис по подгрузке данных о связях в модели
    /// </summary>
    internal static class LoadRLI_Service
    {
        internal static UIControlledApplication RevitUIControlledApp { get; set; }

        internal static void SetStaticEnvironment(UIControlledApplication application) =>
            RevitUIControlledApp = application;

        public static LinkManagerEntity[] CreateLMEntities(UIApplication app)
        {
            List<LinkManagerEntity> result = new List<LinkManagerEntity>();

            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;


            Element[] linkDocColl = uidoc
                .Selection
                .GetElementIds()
                .Select(id => doc.GetElement(id))
                .Where(el => el.GetType() == typeof(RevitLinkType))
                .Cast<RevitLinkType>()
                .ToArray();
            
            if (!linkDocColl.Any()) 
                linkDocColl = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkType))
                    .ToArray();


            ObservableCollection<LinkManagerEntity> linkManagerEntsToUpdate = new ObservableCollection<LinkManagerEntity>();
            foreach (Element linkElem in linkDocColl)
            {
                if (linkElem is RevitLinkType linkType)
                {
                    // Вложенные прикрепленные модели 
                    if (linkType.IsNestedLink)
                        continue;

                    ModelPath newModelPath = null;
                    try
                    {
                        // Анализирую положение связи 
                        ExternalFileReference extFileRef = linkType.GetExternalFileReference();
                        ModelPath oldModelPath = extFileRef.GetAbsolutePath();
                        string oldModelPathString = ModelPathUtils.ConvertModelPathToUserVisiblePath(oldModelPath);
                        if (oldModelPathString.Contains("[*"))
                            oldModelPathString = RemoveSubstringIfExists(oldModelPathString, "[*");

                        if (!CheckWSAvailable(doc, linkType))
                        {
                            result.Add(new LinkManagerUpdateEntity(
                                linkType.get_Parameter(BuiltInParameter.RVT_LINK_FILE_NAME_WITHOUT_EXT).AsString(),
                                oldModelPathString,
                                "РН занят",
                                "Рабочий набор занят, нужно предварительно освободить",
                                EntityStatus.CriticalError));

                            continue;
                        }

                        result.Add(new LinkManagerUpdateEntity(
                            linkType.get_Parameter(BuiltInParameter.RVT_LINK_FILE_NAME_WITHOUT_EXT).AsString(),
                            oldModelPathString,
                            "Не определено",
                            "Не определено",
                            EntityStatus.Error));
                    }
                    catch (FileArgumentNotFoundException)
                    {
                        HtmlOutput.Print($"Файла по переопределенному пути нет. Проверь наличие файла для обновления тут: {ModelPathUtils.ConvertModelPathToUserVisiblePath(newModelPath)}. Если она там есть - обратись к разработчику!", MessageType.Error);
                    }
                }
            }

            return result.OrderBy(ent => ent.LinkName).ToArray();
        }

        /// <summary>
        /// Проверка экземпляров типа на возможность замены
        /// </summary>
        /// <returns>True - свободны, False - заняты</returns>
        public static bool CheckWSAvailable(Document doc, RevitLinkType linkType)
        {
            ElementId[] linkTypeInstanceIds = linkType
                .GetDependentElements(
            new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks))
                .Where(id => doc.GetElement(id) is RevitLinkInstance)
                .ToArray();

            ICollection<ElementId> availableWSElemsId = WorksharingUtils
                .CheckoutElements(doc, linkTypeInstanceIds);

            return linkTypeInstanceIds.Count() == availableWSElemsId.Count();
        }

        /// <summary>
        /// Получить альтернативную модель по указанному пути
        /// </summary>
        public static LinkManagerUpdateEntity GetSimilarByPath(LinkManagerUpdateEntity sourceEnt, string pathToSearch, int revitVersion)
        {
            string sourceName = sourceEnt.LinkName;
            string[] pathParts = pathToSearch.Split('\\');
            System.IO.DirectoryInfo root = new System.IO.DirectoryInfo(pathToSearch);

            int tempLenght = 1000;
            string resultFileName = string.Empty;
            string resultFilePath = string.Empty;
            // Сревер KPLN
            if (root.Exists)
            {
                FileInfo[] rvtFIs = root
                    .GetFiles()
                    .Where(fi => fi.Extension.Equals(".rvt"))
                    .ToArray();

                if (rvtFIs.Length == 0)
                    return new LinkManagerUpdateEntity(sourceEnt.LinkName, sourceEnt.LinkPath, "Ошибка", "В данной папке нет файлов Revit", EntityStatus.Error);

                foreach (FileInfo fi in rvtFIs)
                {
                    string fileName = fi.Name.TrimEnd(".rvt".ToArray());
                    int damLevDist = DamerauLevenshteinDistance(sourceName, fileName);
                    int maxLenghtName = Math.Max(sourceName.Length, fileName.Length);
                    if (damLevDist < tempLenght && damLevDist < maxLenghtName)
                    {
                        tempLenght = damLevDist;
                        resultFileName = fileName;
                        resultFilePath = fi.FullName;
                    }
                }
            }
            // Revit-Server https://www.nuget.org/packages/RevitServerAPILib
            else if (string.IsNullOrEmpty(pathParts[0]))
            {
                string rsHostName = pathParts[2];
                int pathPartsLenght = pathParts.Length;
                if (rsHostName == null)
                {
                    HtmlOutput.Print($"Ошибка заполнения пути для копирования с Revit-Server: ({pathToSearch}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'", MessageType.Error);
                    return null;
                }

                try
                {
                    RevitServer server = new RevitServer(rsHostName, revitVersion);
                    FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 3));

                    if (folderContents.Models.Count == 0)
                        return new LinkManagerUpdateEntity(sourceEnt.LinkName, sourceEnt.LinkPath, "Ошибка", "В данной папке нет файлов Revit", EntityStatus.Error);

                    foreach (var model in folderContents.Models)
                    {
                        string fileName = model.Name.TrimEnd(".rvt".ToArray());
                        int damLevDist = DamerauLevenshteinDistance(sourceName, fileName);
                        int maxLenghtName = Math.Max(sourceName.Length, fileName.Length);
                        if (damLevDist < tempLenght && damLevDist < maxLenghtName)
                        {
                            tempLenght = damLevDist;
                            resultFileName = fileName;
                            resultFilePath = $"RSN:{pathToSearch}\\{model.Name}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    HtmlOutput.Print($"Ошибка открытия Revit-Server ({pathToSearch}):\n{ex.Message}", MessageType.Error);
                    return null;
                }
            }

            if (string.IsNullOrEmpty(resultFileName) || string.IsNullOrEmpty(resultFilePath))
                return new LinkManagerUpdateEntity(sourceEnt.LinkName, sourceEnt.LinkPath, "Ошибка", "В данной папке нет файлов Revit", EntityStatus.Error);

            return new LinkManagerUpdateEntity(sourceEnt.LinkName, sourceEnt.LinkPath, resultFileName, resultFilePath, EntityStatus.Ok);
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

        /// <summary>
        ///  Расчет расстояния Дамерлоу-Левинштейна
        /// </summary>
        private static int DamerauLevenshteinDistance(string firstText, string secondText)
        {
            var n = firstText.Length + 1;
            var m = secondText.Length + 1;
            var arrayD = new int[n, m];

            for (var i = 0; i < n; i++)
            {
                arrayD[i, 0] = i;
            }

            for (var j = 0; j < m; j++)
            {
                arrayD[0, j] = j;
            }

            for (var i = 1; i < n; i++)
            {
                for (var j = 1; j < m; j++)
                {
                    var cost = firstText[i - 1] == secondText[j - 1] ? 0 : 1;

                    arrayD[i, j] = Minimum(arrayD[i - 1, j] + 1,          // удаление
                                            arrayD[i, j - 1] + 1,         // вставка
                                            arrayD[i - 1, j - 1] + cost); // замена

                    if (i > 1 && j > 1
                        && firstText[i - 1] == secondText[j - 2]
                        && firstText[i - 2] == secondText[j - 1])
                    {
                        arrayD[i, j] = Minimum(arrayD[i, j],
                                           arrayD[i - 2, j - 2] + cost); // перестановка
                    }
                }
            }

            return arrayD[n - 1, m - 1];
        }

        private static int Minimum(int a, int b) => a < b ? a : b;

        private static int Minimum(int a, int b, int c) => (a = a < b ? a : b) < c ? a : c;
    }
}
