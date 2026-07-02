using Autodesk.Revit.DB;

namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    /// <summary>
    /// Настройки экспорта файла в IFC
    /// </summary>
    public class DBIFCConfigData : DBConfigEntity
    {
        private IFCVersion _fileVersion;
        private int _spaceBoundaryLevel;
        private bool _wallAndColumnSplitting;
        private bool _exportBaseQuantities;
        private bool _exportLinks;

        private string _viewName;
        private string _worksetToCloseNamesStartWith;
        private string _ifcDocPostfix;

        public DBIFCConfigData() : base()
        {
        }

        public DBIFCConfigData(string name, string pathFrom, string pathTo) : base(name, pathFrom, pathTo)
        {
        }

        /// <summary>
        /// Версия IFC
        /// </summary>
        public IFCVersion FileVersion
        {
            get => _fileVersion;
            set => SetField(ref _fileVersion, value);
        }

        /// <summary>
        /// Уровень границ пространств: 0 / 1 / 2
        /// </summary>
        public int SpaceBoundaryLevel
        {
            get => _spaceBoundaryLevel;
            set => SetField(ref _spaceBoundaryLevel, value);
        }

        /// <summary>
        /// Делить стены и колонны по уровням
        /// </summary>
        public bool WallAndColumnSplitting
        {
            get => _wallAndColumnSplitting;
            set => SetField(ref _wallAndColumnSplitting, value);
        }

        /// <summary>
        /// Экспортировать базовые количества
        /// </summary>
        public bool ExportBaseQuantities
        {
            get => _exportBaseQuantities;
            set => SetField(ref _exportBaseQuantities, value);
        }

        /// <summary>
        /// Экспортировать связанные файлы
        /// </summary>
        public bool ExportLinks
        {
            get => _exportLinks;
            set => SetField(ref _exportLinks, value);
        }

        /// <summary>
        /// Имя вида для экспорта
        /// </summary>
        public string ViewName
        {
            get => _viewName;
            set => SetField(ref _viewName, value);
        }

        /// <summary>
        /// Имя рабочих наборов, которые нужно закрыть
        /// </summary>
        public string WorksetToCloseNamesStartWith
        {
            get => _worksetToCloseNamesStartWith;
            set => SetField(ref _worksetToCloseNamesStartWith, value);
        }

        /// <summary>
        /// Спец. окончание в имени файла IFC
        /// </summary>
        public string IfcDocPostfix
        {
            get => _ifcDocPostfix;
            set => SetField(ref _ifcDocPostfix, value);
        }

        /// <summary>
        /// Метод для копирования данных из другого экземпляра класса
        /// </summary>
        public DBIFCConfigData MergeWithDBConfigEntity(DBIFCConfigData other)
        {
            FileVersion = other.FileVersion;
            SpaceBoundaryLevel = other.SpaceBoundaryLevel;
            WallAndColumnSplitting = other.WallAndColumnSplitting;
            ExportBaseQuantities = other.ExportBaseQuantities;
            ExportLinks = other.ExportLinks;

            ViewName = other.ViewName;
            WorksetToCloseNamesStartWith = other.WorksetToCloseNamesStartWith;
            IfcDocPostfix = other.IfcDocPostfix;

            return this;
        }
    }
}
