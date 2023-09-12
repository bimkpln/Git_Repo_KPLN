using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Documents в БД
    /// </summary>
    public class DocumentDbService : DbService
    {
        internal DocumentDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }
    }
}
