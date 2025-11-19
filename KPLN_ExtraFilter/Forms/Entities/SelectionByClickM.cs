using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Модель для SelectionByClick
    /// </summary>
    public class SelectionByClickM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private IEnumerable<Element> _userSelElems;
        private Element _userSelElem;
        private bool _sameCategory;
        private bool _sameFamily;
        private bool _sameType;
        private bool _sameWorkset;
        private bool _sameParamData;
        private bool _currentView = true;
        private bool _model;
        private bool _belongGroup;

        private ParamEntity _selectedParam;
        private bool _isWorkshared = false;
        private string _userHelp;

        [JsonConstructor]
        public SelectionByClickM() { }

        public SelectionByClickM(Document doc)
        {
            UserSelDoc = doc;
            IsWorkshared = doc.IsWorkshared || doc.IsDetached;
        }

        public Document UserSelDoc { get; set; }

        public IEnumerable<Element> UserSelElems
        {
            get => _userSelElems;
            set
            {
                _userSelElems = value;
                if (_userSelElems.Count() == 1)
                    UserSelElem = UserSelElems.FirstOrDefault();
                else
                    UserSelElem = null;
            }
        }

        public Element UserSelElem
        {
            get => _userSelElem;
            set
            {
                _userSelElem = value;
                
                UpdateUserHelp();
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(SelectedElemParams));
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одинаковой категории
        /// </summary>
        public bool What_SameCategory
        {
            get => _sameCategory;
            set
            {
                _sameCategory = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одинакового семейства
        /// </summary>
        public bool What_SameFamily
        {
            get => _sameFamily;
            set
            {
                _sameFamily = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одинакового типа
        /// </summary>
        public bool What_SameType
        {
            get => _sameType;
            set
            {
                _sameType = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одного рабочего набора
        /// </summary>
        public bool What_Workset
        {
            get => _sameWorkset;
            set
            {
                _sameWorkset = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одного значения параметра
        /// </summary>
        public bool What_ParameterData
        {
            get => _sameParamData;
            set
            {
                _sameParamData = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// Одного значения параметра
        /// </summary>
        public ParamEntity What_SelectedParam
        {
            get => _selectedParam;
            set
            {
                _selectedParam = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(CanRun));
            }
        }

        /// <summary>
        /// В модели
        /// </summary>
        public bool Where_Model
        {
            get => _model;
            set
            {
                _model = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// На виде
        /// </summary>
        public bool Where_CurrentView
        {
            get => _currentView;
            set
            {
                _currentView = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Исключить элементы групп
        /// </summary>
        public bool Belong_Group
        {
            get => _belongGroup;
            set
            {
                _belongGroup = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Коллекция параметров типа/экземпляра у выбранного эл-та
        /// </summary>
        public ParamEntity[] SelectedElemParams
        {
            get
            {
                if (UserSelElem == null)
                    return null;
                
                IEnumerable<Parameter> elemsParams = ParamWorker.GetParamsFromElems(UserSelDoc, new Element[] { UserSelElem });
                if (elemsParams == null)
                    return null;

                List<ParamEntity> allParamsEntities = new List<ParamEntity>(elemsParams.Count());
                foreach (Parameter param in elemsParams)
                {
                    string paramValue;
                    if (param.StorageType == StorageType.String)
                        paramValue = param.AsString();
                    else
                        paramValue = param.AsValueString();

                    string toolTip = $"Значение: {paramValue}";
                    allParamsEntities.Add(new ParamEntity(param, toolTip));
                }

                return allParamsEntities.OrderBy(pe => pe.CurrentParamName).ToArray();
            }

        }

        /// <summary>
        /// Маркер модели из хранилища
        /// </summary>
        public bool IsWorkshared
        {
            get => _isWorkshared;
            set
            {
                _isWorkshared = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Маркер возможности запуска
        /// </summary>
        public bool CanRun
        {
            get
            {
                bool check = UserSelElem != null
                    && (What_SameCategory
                    || What_SameFamily
                    || What_SameType
                    || What_Workset
                    || (What_ParameterData && What_SelectedParam != null));
                
                UpdateUserHelp();
                return check;
            }
        }

        public string UserHelp 
        { 
            get => _userHelp;
            set
            {
                _userHelp = string.IsNullOrWhiteSpace(value) ? string.Empty : $"ВАЖНО: {value}";
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Устанавливаю значения по умолчанию (для сброса сущности)
        /// </summary>
        /// <returns></returns>
        public SelectionByClickM DropToDefault()
        {
            What_SameCategory = false;
            What_SameFamily = false;
            What_SameType = false;
            What_Workset = false;
            What_ParameterData = false;
            Where_Model = true;
            Where_CurrentView = false;
            Belong_Group = false;
            What_SelectedParam = null;

            return this;
        }

        public object ToJson() => new
        {
            // Что выбираем
            this.What_SameCategory,
            this.What_SameFamily,
            this.What_SameType,
            this.What_Workset,
            // Блок для параметров
            this.What_ParameterData,
            this.What_SelectedParam,
            // Где выбираем
            this.Where_Model,
            this.Where_CurrentView,
            // Исключения
            this.Belong_Group,
        };

        protected void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void UpdateUserHelp()
        {
            if (UserSelElem == null)
            {
                UserHelp = "Выбери 1 элемент для поиска аналогов";
                return;
            }

            if (!What_SameCategory &&
                !What_SameFamily &&
                !What_SameType &&
                !What_Workset &&
                !(What_ParameterData))
            {
                UserHelp = "Выбери критерии поиска (блок «Что выбираем»)";
                return;
            }

            if (What_ParameterData && What_SelectedParam == null)
            {
                UserHelp = "Выбери параметр из списка";
                return;
            }

            UserHelp = string.Empty;
        }
    }
}
