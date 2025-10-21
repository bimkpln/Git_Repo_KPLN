using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Creator для таблицы DBProjectsIOSClashMatrix
    /// </summary>
    public class CreatorProjectIOSClashMatrixDbService : AbsCreatorDbService
    {
        public override DbService CreateService()
        {
            SQLFilesExistCheckerAndDBDataSetter();
            string connectionString = CreateConnectionString();

            return new ProjectsIOSClashMatrixDbService(connectionString, DBProjectsIOSClashMatrix.CurrentDB);
        }
    }
}
