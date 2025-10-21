using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Creator для таблицы Documents
    /// </summary>
    public class CreatorDocumentDbService : AbsCreatorDbService
    {
        public override DbService CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new DocumentDbService(connectionString, DBDocument.CurrentDB);
        }
    }
}
