using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame
{
    public class SetParamsByFrameEntity : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _userInputParamValue;
        private string _runButtonName;
        private string _runButtonTooltip;

        [JsonConstructor]
        public SetParamsByFrameEntity()
        {
            RunButtonContext();
        }

        public SetParamsByFrameEntity(IEnumerable<Element> elems, IEnumerable<ParamEntity> paramsEntities) : this()
        {
            SelectedElems = elems.ToArray();
            AllParamEntities = paramsEntities.ToArray();
        }

        public SetParamsByFrameEntity(IEnumerable<Element> elems, IEnumerable<ParamEntity> paramsEntities, IEnumerable<MainItem> mainItems) : this(elems, paramsEntities)
        {
            // Если такой пар-р есть в общей коллекции, значит добавляю его в форму. Иначе - нет
            IEnumerable<MainItem> userSelectedViewModels = mainItems
                .Where(entity => AllParamEntities.Count(ent => ent.RevitParamIntId == entity.UserSelectedParamEntity.RevitParamIntId) == 1)
                .Select(entity =>
                {
                    entity.UserSelectedParamEntity.RevitParamName = AllParamEntities
                    .First(ent => ent.RevitParamIntId == entity.UserSelectedParamEntity.RevitParamIntId)
                    .RevitParamName;

                    return entity;
                });


            MainItems = new ObservableCollection<MainItem>(userSelectedViewModels);
        }

        /// <summary>
        /// Выбранные элементы для анализа
        /// </summary>
        public Element[] SelectedElems { get; private set; }

        /// <summary>
        /// Коллекция ВСЕХ сущностей параметров, которые есть у эл-в
        /// </summary>
        public ParamEntity[] AllParamEntities { get; private set; }

        /// <summary>
        /// Коллекция основных сущностей
        /// </summary>
        public ObservableCollection<MainItem> MainItems { get; private set; } = new ObservableCollection<MainItem>();

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

        public object ToJson() => new
        {
            this.MainItems,
        };

        /// <summary>
        /// Установить данные по кнопке
        /// </summary>
        public void RunButtonContext()
        {
            if (MainItems.Count > 0)
            {
                RunButtonName = "Заполнить!";
                RunButtonTooltip = "Заполнить указанные параметры указанными значениями для выделенных элементов";
            }
            else
            {
                RunButtonName = "Выделить!";
                RunButtonTooltip = "Выделить все вложенности у выбранных элементов";
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
