using Autodesk.Revit.DB;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class DocumentSnapshot
    {
        public string Title { get; private set; }
        public string Path { get; private set; }

        public static DocumentSnapshot FromDocument(Document document)
        {
            return new DocumentSnapshot
            {
                Title = SafeRead(() => document.Title),
                Path = GetDocumentPath(document)
            };
        }

        public string GetBestPath()
        {
            return Path ?? string.Empty;
        }

        private static string GetDocumentPath(Document document)
        {
            string kplnPath = SafeRead(() => KPLN_Library_DBWorker.FactoryParts.SQLite.SQLiteDocService.GetFileFullName(document));
            if (!string.IsNullOrWhiteSpace(kplnPath))
                return kplnPath;

            string centralPath = SafeRead(() =>
            {
                if (!document.IsWorkshared)
                    return string.Empty;

                ModelPath modelPath = document.GetWorksharingCentralModelPath();
                if (modelPath == null)
                    return string.Empty;

                return ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
            });

            if (!string.IsNullOrWhiteSpace(centralPath))
                return centralPath;

            return SafeRead(() => document.PathName);
        }

        private static string SafeRead(System.Func<string> read)
        {
            try
            {
                return read() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}