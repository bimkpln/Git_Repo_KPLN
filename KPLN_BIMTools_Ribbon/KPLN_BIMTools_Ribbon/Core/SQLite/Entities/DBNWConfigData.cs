using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    public class DBNWConfigData : DBConfigEntity
    {
        private double _facetingFactor;
        private bool _convertElementProperties;
        private bool _exportLinks;
        private bool _findMissingMaterials;
        private NavisworksExportScope _exportScope;
        private bool _divideFileIntoLevels;
        private bool _exportRoomGeometry;

        private string _viewName;
        private string _worksetToCloseNamesStartWith;
        private string _navisDocPostfix;

        public DBNWConfigData() : base()
        {
        }

        public DBNWConfigData(string name, string pathFrom, string pathTo) : base(name, pathFrom, pathTo)
        {
        }

        #region Стандартные настройки для Navisworks
        /// <summary>
        /// Коэф. фасетизации
        /// </summary>
        public double FacetingFactor
        {
            get => _facetingFactor;
            set
            {
                SetField(ref _facetingFactor, value);
            }
        }

        /// <summary>
        /// Преобразовать свойства объектов
        /// </summary>
        public bool ConvertElementProperties
        {
            get => _convertElementProperties;
            set
            {
                SetField(ref _convertElementProperties, value);
            }
        }

        /// <summary>
        /// Преобразовать связанные файлы
        /// </summary>
        public bool ExportLinks
        {
            get => _exportLinks;
            set
            {
                SetField(ref _exportLinks, value);
            }
        }

        /// <summary>
        /// Проверять и находить отсутсв. материалы
        /// </summary>
        public bool FindMissingMaterials
        {
            get => _findMissingMaterials;
            set
            {
                SetField(ref _findMissingMaterials, value);
            }
        }

        /// <summary>
        /// Область экспорта (файл, вид, выбранные эл-ты)
        /// </summary>
        public NavisworksExportScope ExportScope
        {
            get => _exportScope;
            set
            {
                SetField(ref _exportScope, value);
            }
        }

        /// <summary>
        /// Разделять файл по уровням
        /// </summary>
        public bool DivideFileIntoLevels
        {
            get => _divideFileIntoLevels;
            set
            {
                SetField(ref _divideFileIntoLevels, value);
            }
        }

        /// <summary>
        /// Экспортировать геомтерию помещений
        /// </summary>
        public bool ExportRoomGeometry
        {
            get => _exportRoomGeometry;
            set
            {
                SetField(ref _exportRoomGeometry, value);
            }
        }
        #endregion

        #region Дополнительные настройки из конфига
        /// <summary>
        /// Имя вида для экспорта
        /// </summary>
        public string ViewName
        {
            get => _viewName;
            set
            {
                SetField(ref _viewName, value);
            }
        }

        /// <summary>
        /// Имя рабочих наборов, которые нужно закрыть (имя начинается с)
        /// </summary>
        public string WorksetToCloseNamesStartWith
        {
            get => _worksetToCloseNamesStartWith;
            set
            {
                SetField(ref _worksetToCloseNamesStartWith, value);
            }
        }

        /// <summary>
        /// Спец. окончание в имени файла Navis
        /// </summary>
        public string NavisDocPostfix
        {
            get => _navisDocPostfix;
            set
            {
                SetField(ref _navisDocPostfix, value);
            }
        }
        #endregion

        /// <summary>
        /// Метод для копирования данных из другого экземпляра класса
        /// </summary>
        /// <param name="other"></param>
        public DBNWConfigData MergeWithDBConfigEntity(DBNWConfigData other)
        {
            // Копируем поля из класса DBNWConfigData
            this.FacetingFactor = other.FacetingFactor;
            this.ConvertElementProperties = other.ConvertElementProperties;
            this.ExportLinks = other.ExportLinks;
            this.FindMissingMaterials = other.FindMissingMaterials;
            this.ExportScope = other.ExportScope;
            this.DivideFileIntoLevels = other.DivideFileIntoLevels;
            this.ExportRoomGeometry = other.ExportRoomGeometry;
            this.ViewName = other.ViewName;
            this.WorksetToCloseNamesStartWith = other.WorksetToCloseNamesStartWith;
            this.NavisDocPostfix = other.NavisDocPostfix;

            return this;
        }
    }
}
