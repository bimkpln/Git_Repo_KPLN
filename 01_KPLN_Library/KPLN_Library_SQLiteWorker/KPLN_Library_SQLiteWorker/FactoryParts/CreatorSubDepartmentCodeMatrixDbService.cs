using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Creator для таблицы Documents
    /// </summary>
    public class CreatorSubDepartmentCodeMatrixDbService : AbsCreatorDbService
    {
        public override DbService CreateService()
        {
            SQLFilesExistChecker();
            string connectionString = CreateConnectionString("KPLN_Loader_MainDB");

            return new SubDepartmentCodeMatrixDbService(connectionString, DBSubDepartmentCodeMatrix.CurrentDB);
        }
    }
}
