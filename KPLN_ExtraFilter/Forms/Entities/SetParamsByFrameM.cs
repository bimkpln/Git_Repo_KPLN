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
    /// <summary>
    /// Сценарий запуска
    /// </summary>
    public enum SetParamsByFrameScript
    {
        /// <summary>
        /// Выделить элементы, включая вложенности
        /// </summary>
        SelectDependentElements,
        /// <summary>
        /// Запуск рамки выбора
        /// </summary>
        SelectElementsByFrame,
        /// <summary>
        /// Заполнить параметры, включая вложенности
        /// </summary>
        SetElementsParameters,
    }


    public class SetParamsByFrameM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private View _docActiveView;
        private IEnumerable<Element> _userSelElems;
        private string _userInputParamValue;
        private string _runButtonName;
        private string _runButtonTooltip;
        private bool _canAddParam = false;
        private bool _canRun = false;
        private string _userHelp;

        [JsonConstructor]
        public SetParamsByFrameM()
        {
            SetScriptAndRunButtonContext();
        }

        public SetParamsByFrameM(Document doc)
        {
            Doc = doc;
            UpdateScriptANDCanRunANDUserHelp();
        }

        public SetParamsByFrameM(Document doc, IEnumerable<Element> userSelElems)
        {
            Doc = doc;
            UserSelElems = userSelElems;
            UpdateScriptANDCanRunANDUserHelp();
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
        /// Тип текущего сценария
        /// </summary>
        public SetParamsByFrameScript CurrentScript { get; private set; } = SetParamsByFrameScript.SelectElementsByFrame;

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

                    UpdateScriptANDCanRunANDUserHelp();
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
        /// Маркер возможности добавления параметра
        /// </summary>
        public bool CanAddParam
        {
            get => _canAddParam;
            set
            {
                _canAddParam = value;
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
        /// Обновить статус возможности запуска и текстовых подсказок
        /// </summary>
        public void UpdateScriptANDCanRunANDUserHelp()
        {
            SetScriptAndRunButtonContext();

            // Не выделен элемент
            if (UserSelElems == null || UserSelElems.Count() == 0)
            {
                UserHelp = "Выбери элементы для анализа";
                CanAddParam = false;
                CanRun = true;
                return;
            }

            // Не выбран параметр
            bool isParamItem = ParamItems.Any(paramM => paramM.ParamM_SelectedParameter == null);
            if (isParamItem)
            {
                UserHelp = "Выбери параметр для заполнения, или удали (кнопака Очистить, или через ПКМ по параметру)";
                CanAddParam = false;
                CanRun = false;
                return;
            }

            // Не указано значение в параметре
            if (!isParamItem && ParamItems.Any(paramM => paramM.ParamM_InputValue == null))
            {
                UserHelp = "Укажи значение параметра для заполнения, или удали (кнопака Очистить, или через ПКМ по параметру)";
                CanAddParam = false;
                CanRun = false;
                return;
            }

            UserHelp = string.Empty;
            CanAddParam = true;
            CanRun = true;
        }

        /// <summary>
        /// Определиться со сценарием запуска и установить данные по кнопке
        /// </summary>
        private void SetScriptAndRunButtonContext()
        {
            // Обновляю сценарий
            if (ParamItems.Count > 0 && UserSelElems.Count() != 0)
                CurrentScript = SetParamsByFrameScript.SetElementsParameters;
            else if (ParamItems.Count > 0 && UserSelElems.Count() == 0)
                CurrentScript = SetParamsByFrameScript.SelectElementsByFrame;
            else if (UserSelElems.Count() == 0)
                CurrentScript = SetParamsByFrameScript.SelectElementsByFrame;
            else
                CurrentScript = SetParamsByFrameScript.SelectDependentElements;


            // Устанавливаю кнопки
            switch (CurrentScript)
            {
                case SetParamsByFrameScript.SetElementsParameters:
                    RunButtonName = "Заполнить!";
                    RunButtonTooltip = "Заполнить указанные параметры указанными значениями для выделенных элементов, включая вложенные";
                    break;
                case SetParamsByFrameScript.SelectElementsByFrame:
                    RunButtonName = "Выделить рамкой!";
                    RunButtonTooltip = "Включается режим рамки, которая может выбирать элементы не реагируя на вложенность в группу";
                    break;
                case SetParamsByFrameScript.SelectDependentElements:
                    RunButtonName = "Выделить!";
                    RunButtonTooltip = "К выделенным в модели элементам - добавятся вложенные элементы";
                    break;
            }
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
