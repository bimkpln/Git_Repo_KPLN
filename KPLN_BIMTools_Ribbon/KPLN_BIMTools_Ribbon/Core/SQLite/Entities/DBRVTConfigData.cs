namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    /// <summary>
    /// Настройки экспорта файла на RevitServer
    /// </summary>
    public class DBRVTConfigData : DBConfigEntity
    {
        private int _maxBackup = -1;
        private string _nameChangeFind = "🔐";
        private string _nameChangeSet = "🔐";

        public DBRVTConfigData() : base()
        {
        }

        public DBRVTConfigData(string name, string pathFrom, string pathTo) : base(name, pathFrom, pathTo)
        {
        }

        #region Стандартные настройки для сохранения rvt-файла
        /// <summary>
        /// Колчество бэкапов
        /// </summary>
        public int MaxBackup
        {
            get => _maxBackup;
            set
            {
                SetField(ref _maxBackup, value);
            }
        }

        /// <summary>
        /// Замена: Часть имени для поиска замены
        /// </summary>
        public string NameChangeFind
        {
            get => _nameChangeFind;
            set
            {
                SetField(ref _nameChangeFind, value);
            }
        }

        /// <summary>
        /// Замена: Новая часть имени файла
        /// </summary>
        public string NameChangeSet
        {
            get => _nameChangeSet;
            set
            {
                SetField(ref _nameChangeSet, value);
            }
        }
        #endregion


        /// <summary>
        /// Метод для копирования данных из другого экземпляра класса
        /// </summary>
        /// <param name="other"></param>
        public DBRVTConfigData MergeWithDBConfigEntity(DBRVTConfigData other)
        {
            // Копируем поля из класса DBRVTConfigData
            this.NameChangeFind = other.NameChangeFind;
            this.NameChangeSet = other.NameChangeSet;
            this.MaxBackup = other.MaxBackup;

            return this;
        }
    }
}
