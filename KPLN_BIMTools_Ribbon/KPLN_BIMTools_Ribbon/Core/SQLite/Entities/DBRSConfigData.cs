namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    /// <summary>
    /// Настройки экспорта файла на RevitServer
    /// </summary>
    public class DBRSConfigData : DBConfigEntity
    {
        public DBRSConfigData() : base()
        {
        }

        public DBRSConfigData(string name, string pathFrom, string pathTo) : base(name, pathFrom, pathTo)
        {
        }
    }
}
