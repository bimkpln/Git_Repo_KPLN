using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Data.SQLite;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Creator для таблицы Users
    /// </summary>>
    public class SQLiteCreatorUserDbService : AbsCreatorDbService<SQLiteConnection, SQLiteException>
    {
        public override DbServiceAbstr<SQLiteConnection, SQLiteException> CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new SQLiteUserService(connectionString, DBUser.CurrentDB);
        }
    }
}
