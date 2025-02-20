namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    /// <summary>
    /// Настройки экспорта файла на RevitServer
    /// </summary>
    public class DBRVTConfigData : DBConfigEntity
    {
        private int _maxBackup = -1;

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
        #endregion


        /// <summary>
        /// Метод для копирования данных из другого экземпляра класса
        /// </summary>
        /// <param name="other"></param>
        public DBRVTConfigData MergeWithDBConfigEntity(DBRVTConfigData other)
        {
            // Копируем поля из класса DBRVTConfigData
            this.MaxBackup = other.MaxBackup;

            return this;
        }
    }
}
