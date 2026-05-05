using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using MySqlConnector;

namespace KPLN_Library_DBWorker.FactoryParts.MySql
{
    /// <summary>
    /// Creator для таблицы Users
    /// </summary>>
    public class CreatorUserDbService : AbsCreatorDbService<MySqlConnection, MySqlException>
    {
        public override DbServiceAbstr<MySqlConnection, MySqlException> CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new UserDbService(connectionString, DBUser.CurrentDB);
        }
    }
}
