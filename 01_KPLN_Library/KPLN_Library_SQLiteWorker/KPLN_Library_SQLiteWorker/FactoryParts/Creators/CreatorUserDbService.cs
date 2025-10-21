using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Creator для таблицы Users
    /// </summary>>
    public class CreatorUserDbService : AbsCreatorDbService
    {
        public override DbService CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new UserDbService(connectionString, DBUser.CurrentDB);
        }
    }
}
