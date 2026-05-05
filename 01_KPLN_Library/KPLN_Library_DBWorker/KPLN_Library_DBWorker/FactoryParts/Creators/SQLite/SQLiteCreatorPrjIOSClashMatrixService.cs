using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Data.SQLite;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Creator для таблицы DBProjectsIOSClashMatrix
    /// </summary>
    public class SQLiteCreatorPrjIOSClashMatrixService : AbsCreatorDbService<SQLiteConnection, SQLiteException>
    {
        public override DbServiceAbstr<SQLiteConnection, SQLiteException> CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new SQLitePrjOSClashMatrixService(connectionString, DBProjectsIOSClashMatrix.CurrentDB);
        }
    }
}
