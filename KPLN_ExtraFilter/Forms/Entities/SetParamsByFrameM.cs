using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    public class SetParamsByFrameM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private View _docActiveView;
        private IEnumerable<Element> _userSelElems;
        private string _userInputParamValue;
        private string _runButtonName;
        private string _runButtonTooltip;
        private bool _canRun = false;
        private string _userHelp;

        [JsonConstructor]
        public SetParamsByFrameM()
        {
            RunButtonContext();
        }

        public SetParamsByFrameM(Document doc)
        {
            Doc = doc;
            UpdateCanRunANDUserHelp();
        }

        public SetParamsByFrameM(Document doc, IEnumerable<Element> userSelElems)
        {
            Doc = doc;
            UserSelElems = userSelElems;
            UpdateCanRunANDUserHelp();
        }


        public Document Doc { get; set; }

        public View DocActiveView
        {
            get => _docActiveView;
            set
            {
                if (value != _docActiveView)
                {
                    _docActiveView = value;
                }
            }
        }

        /// <summary>
        /// Пользовательский выбор из модели
        /// </summary>
        public IEnumerable<Element> UserSelElems
        {
            get => _userSelElems;
            set
            {
                if (value != _docActiveView)
                {
                    _userSelElems = value;
                    if (ParamItems.Count > 0)
                        ReloadParams();

                    UpdateCanRunANDUserHelp();
                }
            }
        }

        /// <summary>
        /// Коллекция основных сущностей
        /// </summary>
        public ObservableCollection<SetParamsByFrameM_ParamM> ParamItems { get; set; } = new ObservableCollection<SetParamsByFrameM_ParamM>();

        /// <summary>
        /// Имя кнопки запуска (в зависимости от логики)
        /// </summary>
        public string RunButtonName
        {
            get => _runButtonName;
            private set
            {
                _runButtonName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Описание кнопки запуска (в зависимости от логики)
        /// </summary>
        public string RunButtonTooltip
        {
            get => _runButtonTooltip;
            private set
            {
                _runButtonTooltip = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Введенное пользователем значение параметра
        /// </summary>
        public string UserInputParamValue
        {
            get => _userInputParamValue;
            set
            {
                _userInputParamValue = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Маркер возможности запуска
        /// </summary>
        public bool CanRun
        {
            get => _canRun;
            set
            {
                _canRun = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Пользовательская подсказка
        /// </summary>
        public string UserHelp
        {
            get => _userHelp;
            set
            {
                _userHelp = string.IsNullOrWhiteSpace(value) ? string.Empty : $"ВАЖНО: {value}";
                NotifyPropertyChanged();
            }
        }

        public object ToJson() => new
        {
            this.ParamItems,
        };

        /// <summary>
        /// Установить данные по кнопке
        /// </summary>
        public void RunButtonContext()
        {
            if (ParamItems.Count > 0)
            {
                RunButtonName = "Заполнить!";
                RunButtonTooltip = "Заполнить указанные параметры указанными значениями для выделенных элементов, включая вложенные";
            }
            else
            {
                RunButtonName = "Выделить!";
                RunButtonTooltip = "К выделенным в модели элементам - добавятся вложенные элементы";
            }
        }

        /// <summary>
        /// Обновить статус возможности запуска и текстовых подсказок
        /// </summary>
        public void UpdateCanRunANDUserHelp()
        {
            // Не выделен элемент
            if (UserSelElems == null || UserSelElems.Count() == 0)
            {
                UserHelp = "Выбери элементы для анализа";
                CanRun = false;
                return;
            }

            // Не выбран параметр
            bool isParamItem = ParamItems.Any(paramM => paramM.ParamM_SelectedParameter == null);
            if (isParamItem)
            {
                UserHelp = "Выбери параметр для заполнения";
                CanRun = false;
                return;
            }

            // Не указано значение в параметре
            if (!isParamItem && ParamItems.Any(paramM => paramM.ParamM_InputValue == null))
            {
                UserHelp = "Укажи значение параметра для заполнения";
                CanRun = false;
                return;
            }

            UserHelp = string.Empty;
            CanRun = true;
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>
        /// Перезагружает список параметров
        /// </summary>
        private void ReloadParams()
        {
            foreach (var paramM in ParamItems)
            {
                var tempOldSelParamM = paramM.ParamM_SelectedParameter;
                paramM.ParamM_UserSelElems = UserSelElems;

                if (tempOldSelParamM != null)
                    paramM.RestoreSelectedParamById(tempOldSelParamM.RevitParamIntId);
            }
        }
    }
}
