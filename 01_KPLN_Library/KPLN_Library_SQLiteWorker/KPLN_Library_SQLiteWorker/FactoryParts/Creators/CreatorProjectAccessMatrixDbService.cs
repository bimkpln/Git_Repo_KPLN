using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Creator для таблицы DBProjectsAccessMatrix
    /// </summary>
    public class CreatorProjectAccessMatrixDbService : AbsCreatorDbService
    {
        public override DbService CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new ProjectsAccessMatrixDbService(connectionString, DBProjectsAccessMatrix.CurrentDB);
        }
    }
}
