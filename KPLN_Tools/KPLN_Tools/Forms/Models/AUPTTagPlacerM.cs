using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models
{
    public sealed class AUPTTagPlacerM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _selectedTagTypeName;
        private string _auptSystemTypeNameMainPart = "АУПТ";
        private int _minPipeLenght = 500;
        private bool _ignoreTagged = true;
        private bool _ingoreMainPipe = true;

        private Document _doc;
        private FamilySymbol _selectedTagType;

        [JsonConstructor]
        public AUPTTagPlacerM() { }

        /// <summary>
        /// Доступные имена типов марки из модели
        /// </summary>
        public string[] DocTagTypeNames { get; private set; }

        /// <summary>
        /// Имя выбранного типа марки
        /// </summary>
        public string SelectedTagTypeName
        {
            get => _selectedTagTypeName;
            set
            {
                _selectedTagTypeName = value;
                OnPropertyChanged();

                CheckAndSetSelectedTagType();
            }
        }

        /// <summary>
        /// Ключевой фрагмент имени системы АУПТ
        /// </summary>
        public string AUPTSystemTypeNameMainPart
        {
            get => _auptSystemTypeNameMainPart;
            set
            {
                _auptSystemTypeNameMainPart = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Выбранный минимальный размер трубы
        /// </summary>
        public int MINPipeLenght

        {
            get => _minPipeLenght;
            set
            {
                if (_minPipeLenght != value)
                {
                    _minPipeLenght = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Игнорировать при маркировке уже промаркированные?
        /// </summary>
        public bool IgnoreTagged
        {
            get => _ignoreTagged;
            set
            {
                _ignoreTagged = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Игнорировать при маркировке магистраль?
        /// </summary>
        public bool IngoreMainPipe
        {
            get => _ingoreMainPipe;
            set
            {
                _ingoreMainPipe = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Выбранный тип марки
        /// </summary>
        public FamilySymbol SelectedTagType
        {
            get => _selectedTagType;
            private set
            {
                _selectedTagType = value;
            }
        }

        public object ToJson() => new
        {
            this.SelectedTagTypeName,
            this.AUPTSystemTypeNameMainPart,
            this.MINPipeLenght,
            this.IgnoreTagged,
            this.IngoreMainPipe,
        };

        /// <summary>
        /// Дополняю сущность данными по документу
        /// </summary>
        /// <returns></returns>
        public AUPTTagPlacerM SetMainData(Document doc)
        {
            _doc = doc;

            DocTagTypeNames = GetPipeTags(doc)
                .Select(fs => fs.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString())
                .OrderBy(n => n)
                .ToArray();

            CheckAndSetSelectedTagType();

            return this;
        }

        private void CheckAndSetSelectedTagType()
        {
            if (string.IsNullOrEmpty(SelectedTagTypeName) || _doc == null)
                return;

            FamilySymbol tagSymbol = GetPipeTags(_doc).FirstOrDefault(fs =>
                fs
                .get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM)
                .AsString()
                .Equals(SelectedTagTypeName));

            if (tagSymbol == null)
            {
                SelectedTagType = null;
                SelectedTagTypeName = string.Empty;
            }
            else
                SelectedTagType = tagSymbol;
        }

        private void OnPropertyChanged([CallerMemberName] string propName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));

        private IEnumerable<FamilySymbol> GetPipeTags(Document doc) =>
            new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_PipeTags)
                .Cast<FamilySymbol>();
    }
}
