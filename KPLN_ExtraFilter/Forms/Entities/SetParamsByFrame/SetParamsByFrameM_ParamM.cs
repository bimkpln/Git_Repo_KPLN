using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame
{
    /// <summary>
    /// Класс-сущность для WPF окна. Им комплектуется ItemsControl
    /// </summary>
    public sealed class SetParamsByFrameM_ParamM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly SetParamsByFrameM _modelM;
        private IEnumerable<Element> _paramM_UserSelElems;
        private ParamEntity _paramM_SelectedParameter;
        private string _paramM_InputValue;

        [JsonConstructor]
        public SetParamsByFrameM_ParamM()
        {
        }

        public SetParamsByFrameM_ParamM(SetParamsByFrameM modelM)
        {
            _modelM = modelM;

            ParamM_Doc = modelM.Doc;
            ParamM_UserSelElems = modelM.UserSelElems.ToArray();
        }

        public Document ParamM_Doc { get; set; }

        public IEnumerable<Element> ParamM_UserSelElems
        {
            get => _paramM_UserSelElems;
            set
            {
                // запомним имя предыдущего выбора (если был)
                int prevParamId = _paramM_SelectedParameter == null ? -1 : _paramM_SelectedParameter.RevitParamIntId;

                _paramM_UserSelElems = value;
                NotifyPropertyChanged();

                ParamM_ParamFilter.LoadCollection(ParamM_DocParameters);
                ParamM_ParamFilter.View?.Refresh();
                RestoreSelectedParamById(prevParamId);

                NotifyPropertyChanged(nameof(ParamM_FilteredParamView));
            }
        }

        /// <summary>
        /// Экземпляр фильтра по параметру
        /// </summary>
        public CollectionFilter<ParamEntity> ParamM_ParamFilter { get; } = new CollectionFilter<ParamEntity>();

        /// <summary>
        /// Пользовательский ввод для поиска параметра
        /// </summary>
        public string SearchParamText
        {
            get => ParamM_ParamFilter.SearchText;
            set => ParamM_ParamFilter.SearchText = value;
        }

        public ICollectionView ParamM_FilteredParamView => ParamM_ParamFilter.View;

        /// <summary>
        /// Коллекция параметров типа/экземпляра у выбранного эл-та
        /// </summary>
        public ParamEntity[] ParamM_DocParameters
        {
            get
            {
                IEnumerable<Parameter> elemsParams = DocWorker.GetAllParamsFromElems(ParamM_Doc, ParamM_UserSelElems, true);
                if (elemsParams == null)
                    return null;

                List<ParamEntity> allParamsEntities = new List<ParamEntity>(elemsParams.Count());
                foreach (Parameter param in elemsParams)
                {
                    string toolTip = string.Empty;
                    if (param.IsShared)
                        toolTip = $"Id: {param.Id}, GUID: {param.GUID}";
                    else if (param.Id.IntegerValue < 0)
                        toolTip = $"Id: {param.Id}, это СИСТЕМНЫЙ параметр проекта";
                    else
                        toolTip = $"Id: {param.Id}, это ПОЛЬЗОВАТЕЛЬСКИЙ параметр проекта";

                    allParamsEntities.Add(new ParamEntity(param, toolTip));
                }


                return allParamsEntities.OrderBy(pe => pe.RevitParamName).ToArray();
            }
        }

        /// <summary>
        /// Выбранный пользователем параметр
        /// </summary>
        public ParamEntity ParamM_SelectedParameter
        {
            get => _paramM_SelectedParameter;
            set
            {
                _paramM_SelectedParameter = value;
                NotifyPropertyChanged();

                _modelM.UpdateCanRunANDUserHelp();
                _modelM.RunButtonContext();
            }
        }

        /// <summary>
        /// Введенное пользователем значение
        /// </summary>
        public string ParamM_InputValue
        {
            get => _paramM_InputValue;
            set
            {
                _paramM_InputValue = value;
                NotifyPropertyChanged();

                _modelM.UpdateCanRunANDUserHelp();
            }
        }

        public object ToJson() => new
        {
            // Parameter - не стоит добавлять в JSON, переваривается плохо.
            // Нужно на чтении JSON уточнять значения по CurrentParam
            this.ParamM_SelectedParameter,
            this.ParamM_InputValue,
        };

        /// <summary>
        /// Восстанавливает выбранный параметр, если в текущем View есть параметр с тем же id.
        /// </summary>
        public void RestoreSelectedParamById(int prevParamId)
        {
            // если нет предыдущего выбора или View — просто ничего не делаем
            if (prevParamId == -1 || ParamM_ParamFilter?.View == null)
                return;

            // Проходим по элементам View и ищем экземпляр с тем же именем
            var match = ParamM_ParamFilter
                .View
                .Cast<ParamEntity>()
                .FirstOrDefault(c => c.RevitParamIntId == prevParamId);

            // присваиваем существующий экземпляр из коллекции — тогда ComboBox его увидит
            if (match != null)
                ParamM_SelectedParameter = match;
            // опционально: снять выбор, если совпадения нет
            else
                ParamM_SelectedParameter = null;
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
