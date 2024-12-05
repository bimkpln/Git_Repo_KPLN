using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms.Entities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Entities
{
    public class SelectionByClickEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _sameCategory;
        private bool _sameFamily;
        private bool _sameType;
        private bool _sameWorkset;
        private bool _sameParamData;
        private bool _model;
        private bool _currentView;
        private bool _belongGroup;

        private ParamEntity _selectedParam;

        public SelectionByClickEntity(Document doc, Element userSelElem)
        {
            Parameter[] elemsParams = ParamWorker.GetParamsFromElems(doc, new Element[] { userSelElem }).ToArray();
            
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

            SelectedElemParams = allParamsEntities.OrderBy(pe => pe.CurrentParamName).ToArray();
            if (SelectedElemParams.Any())
                What_SelectedParam = SelectedElemParams.FirstOrDefault();
        }

        /// <summary>
        /// Одинаковой категории
        /// </summary>
        public bool What_SameCategory
        {
            get => _sameCategory;
            set
            {
                if (_sameCategory != value)
                {
                    _sameCategory = value;
                    NotifyPropertyChanged();
                }
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
                if (_sameFamily != value)
                {
                    _sameFamily = value;
                    NotifyPropertyChanged();
                }
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
                if (_sameType != value)
                {
                    _sameType = value;
                    NotifyPropertyChanged();
                }
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
                if (_sameWorkset != value)
                {
                    _sameWorkset = value;
                    NotifyPropertyChanged();
                }
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
                if (_sameParamData != value)
                {
                    _sameParamData = value;
                    NotifyPropertyChanged();
                }
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
                if (_selectedParam != value)
                {
                    _selectedParam = value;
                    NotifyPropertyChanged();
                }
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
                if (_model != value)
                {
                    _model = value;
                    NotifyPropertyChanged();
                }
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
                if (_currentView != value)
                {
                    _currentView = value;
                    NotifyPropertyChanged();
                }
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
                if (_belongGroup != value)
                {
                    _belongGroup = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коллекция параметров типа/экземпляра у выбранного эл-та
        /// </summary>
        public ParamEntity[] SelectedElemParams { get; private set; }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
