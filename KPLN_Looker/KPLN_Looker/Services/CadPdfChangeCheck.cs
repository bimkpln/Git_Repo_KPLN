using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker.Core;
using KPLN_Looker.ExecutableCommand;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KPLN_Looker.Services
{
    internal static class CadPdfChangeCheck
    {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);

        private static readonly string[] _transactionNames =
        {
            "импорт векторных данных", "import vector data",
            "импорт изображения", "import image", "импорт pdf", "import pdf",
            "связать сапр", "link cad", "связать pdf", "link pdf"
        };

        internal static bool IsRelevantTransaction(DocumentChangedEventArgs args)
        {
            // Undo не проверяем повторно; Redo запрещённой вставки снова отменяем.
            if (args.Operation != UndoOperation.TransactionCommitted && args.Operation != UndoOperation.TransactionRedone)
                return false;

            return args.GetTransactionNames().Any(IsRelevantTransactionName);
        }

        internal static bool IsRelevantTransactionName(string name) =>
            name != null && _transactionNames.Any(part => name.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0);

        internal static bool Check(Document doc, DocumentChangedEventArgs args, DBProject project)
        {
            var errors = new List<string>();
            var summaries = new List<string>();
            var checkedTypes = new HashSet<ElementId>();

            // Проверяем только затронутые элементы, включая перезагрузку и замену типа связи.
            foreach (ElementId id in args.GetAddedElementIds().Concat(args.GetModifiedElementIds()).Distinct())
            {
                Element element = doc.GetElement(id);
                Element type = element is ImportInstance || element is ImageInstance
                    ? doc.GetElement(element.GetTypeId())
                    : element;

                if (type == null || checkedTypes.Contains(type.Id))
                    continue;

                try
                {
                    string path;
                    bool linked;
                    string name = type.Name;

                    if (element is ImportInstance cad)
                    {
                        linked = cad.IsLinked;
                        path = linked ? GetLinkPath(type) : name;
                        // Импорт CAD может иметь имя без расширения: определяем импорт только по IsLinked.
                        if (linked && !HasExtension(path, ".dwg"))
                            continue;
                    }
                    else if (type is CADLinkType && type.IsExternalFileReference())
                    {
                        path = GetLinkPath(type);
                        linked = true;
                        if (!HasExtension(path, ".dwg"))
                            continue;
                    }
                    else if (type is ImageType image)
                    {
                        path = image.Path;
                        if (!HasExtension(path, ".pdf"))
                            continue;

#if Debug2020 || Revit2020
                        // В Revit 2020 PDF поддерживается только как импорт.
                        linked = false;
#else
                        linked = image.Source == ImageTypeSource.Link;
                        if (linked)
                            path = GetImageLinkPath(doc, image);
#endif
                    }
                    else
                        continue;

                    checkedTypes.Add(type.Id);

                    if (!linked)
                    {
                        summaries.Add("Импорт CAD/PDF запрещён. Используйте только связь.");
                        errors.Add($"Файл: {name}\n\nПричина: файл вставлен как импорт.\n\nЧто делать: используйте команду связывания файла.");
                        continue;
                    }

                    // MainPath — папка на файловом сервере и для проектов с моделями на Revit Server.
                    string root = NormalizePath(project?.MainPath);
                    string linkPath = NormalizePath(path);
                    if (root == null)
                    {
                        summaries.Add("Не удалось определить папку проекта для проверки связи. Обратитесь в BIM-отдел.");
                        errors.Add($"Файл: {name}\n\nПричина: в БД Projects не определена папка проекта MainPath.\n\nЧто делать: обратитесь в BIM-отдел.");
                    }
                    else if (linkPath == null || !linkPath.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase))
                    {
                        summaries.Add("Вы вставляете связь не из папки проекта. Это запрещено.\n\nХраните и используйте файлы в папке проекта по структуре KPLN.");
                        errors.Add($"Файл: {name}\n\nПричина: связь находится вне папки проекта.\n\nТекущий путь:\n{path}\n\nПапка проекта:\n{project.MainPath}\n\nЧто делать: поместите файл в соответствующую папку по структуре KPLN и повторите связывание.");
                    }
                }
                catch (Exception ex)
                {
                    checkedTypes.Add(type.Id);
                    summaries.Add("Не удалось проверить файл CAD/PDF. Обратитесь в BIM-отдел.");
                    errors.Add($"Файл: {type.Name}\n\nОшибка проверки:\n{ex.Message}\n\nЧто делать: обратитесь в BIM-отдел.");
                }
            }

            if (errors.Count == 0)
                return false;

            // DocumentChanged доступен только для чтения: ставим Undo в существующую очередь Idling.
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new UndoEvantHandler(doc));
            new TaskDialog("KPLN: Ошибка DWG/PDF")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainInstruction = "Недопустимая операция DWG/PDF будет отменена",
                MainContent = string.Join("\n\n", summaries.Distinct()),
                ExpandedContent = string.Join("\n\n────────────────────\n\n", errors),
                CommonButtons = TaskDialogCommonButtons.Ok
            }.Show();

            return true;
        }

        private static string GetLinkPath(Element type)
        {
            ExternalFileReference reference = type.GetExternalFileReference();
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(reference.GetAbsolutePath());
        }

        private static string GetImageLinkPath(Document doc, ImageType image)
        {
            if (image.IsExternalFileReference())
                return GetLinkPath(image);

            if (image.PathType == PathType.Absolute)
                return image.Path;

            if (image.PathType == PathType.Relative)
            {
                string modelPath = KPLN_Library_DBWorker.FactoryParts.SQLite.SQLiteDocService.GetFileFullName(doc);
                if (NormalizePath(modelPath) != null)
                    return Path.GetFullPath(Path.Combine(Path.GetDirectoryName(modelPath), image.Path));
            }

            throw new InvalidOperationException("Не удалось определить абсолютный файловый путь PDF-связи. Задайте абсолютный путь к файлу на сервере проекта");
        }

        private static bool HasExtension(string path, string extension) =>
            !string.IsNullOrWhiteSpace(path) && path.EndsWith(extension, StringComparison.OrdinalIgnoreCase);

        internal static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Equals("Empty", StringComparison.OrdinalIgnoreCase))
                return null;

            path = path.Trim().Replace('/', '\\');
            // Не разрешаем относительные пути через текущий рабочий каталог процесса Revit.
            if (!path.StartsWith("\\\\") && !(path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' && path[2] == '\\'))
                return null;

            path = Path.GetFullPath(path).TrimEnd('\\');
            if (path.Length >= 2 && path[1] == ':')
            {
                int length = 32768;
                var remote = new StringBuilder(length);
                if (WNetGetConnection(path.Substring(0, 2), remote, ref length) == 0)
                    path = remote.ToString().TrimEnd('\\') + path.Substring(2);
            }

            return path;
        }
    }
}
