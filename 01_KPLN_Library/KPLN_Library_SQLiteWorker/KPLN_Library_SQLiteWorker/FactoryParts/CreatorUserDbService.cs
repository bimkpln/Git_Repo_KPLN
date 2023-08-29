using KPLN_Library_SQLiteWorker.Abstractions;
using KPLN_Library_SQLiteWorker.FactoryParts.Abstractions;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    public class CreatorUserDbService : AbsCreatorDbService
    {
        public override AbsDbService CreateService()
        {
            SQLFilesExistChecker();
            string connectionString = CreateConnectionString("KPLN_Loader_MainDB");

            return new UserDbService(connectionString);
        }
    }
}
