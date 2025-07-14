using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.Entities.SelectionByClick
{
    public class SelectionByClickEntity : INotifyPropertyChanged, IJsonSerializable
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

        [JsonConstructor]
        public SelectionByClickEntity()
        {
        }

        public SelectionByClickEntity(Document doc, Element userSelElem)
        {
            // Устанавливаю значения по умолчанию (для сброса сущности)
            What_SameCategory = false;
            What_SameFamily = false;
            What_SameType = false;
            What_Workset = false;
            What_ParameterData = false;
            Where_Model = true;
            Where_CurrentView = false;
            Belong_Group = false;
            UpdateParams(doc, userSelElem, true);
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
        public ParamEntity[] SelectedElemParams { get; private set; }

        public void UpdateParams(Document doc, Element userSelElem, bool setDefault = false)
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
            if (SelectedElemParams.Any() && setDefault)
                What_SelectedParam = SelectedElemParams.FirstOrDefault();
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

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    }
}
