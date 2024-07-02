using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using UIFramework;

namespace KPLN_Tools.Common.SS_System
{
    public class SS_SystemViewEntity : INotifyPropertyChanged
    {
        private GraphicsStyle _selectedStyle;
        private ElectricalSystemType _selectedSystemType;
        private bool _isLineDraw = false;
        private string _systemNumber;
        private int _startNumber;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Коллекция типов линий в проекте
        /// </summary>
        public Dictionary<string, GraphicsStyle> LineStyles {  get; set; }

        /// <summary>
        /// Коллекция типов систем СС
        /// </summary>
        public Dictionary<string, ElectricalSystemType> ElectricalSystemTypes { get; set; }

        /// <summary>
        /// Выбранный тип линии для построения в проекте
        /// </summary>
        public GraphicsStyle SelectedStyle 
        {
            get => _selectedStyle;
            set
            {
                if (_selectedStyle != value)
                {
                    _selectedStyle = value;
                    OnPropertyChanged();
                }
            } 
        }

        /// <summary>
        /// Выбранный тип системы для построения в проекте
        /// </summary>
        public ElectricalSystemType SelectedSystemType
        {
            get => _selectedSystemType;
            set
            {
                if (_selectedSystemType != value)
                {
                    _selectedSystemType = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Строить линию в проекте?
        /// </summary>
        public bool IsLineDraw
        {
            get => _isLineDraw;
            set
            {
                if (_isLineDraw != value)
                {
                    _isLineDraw = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Номер цепи
        /// </summary>
        public string SystemNumber 
        {
            get => _systemNumber;
            set 
            {
                if (_systemNumber != value)
                {
                    _systemNumber = value;
                    OnPropertyChanged();

                    string[] splitedNumber = SystemNumberSplit();
                    if (!int.TryParse(splitedNumber[0], out int elemIndex))
                        throw new Exception("Скинь разработчику: Не удалось преобразовать номер в тип int.");
                    else
                        StartNumber = elemIndex;

                    UserSystemIndex = splitedNumber[1];
                    UserSeparator = splitedNumber[2];
                }
            }
        }

        /// <summary>
        /// Стартовый номер цепи
        /// </summary>
        public int StartNumber 
        {
            get => _startNumber;
            set
            {
                if (_startNumber != value)
                {
                    _startNumber = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Пользовательский индекс системы
        /// </summary>
        public string UserSystemIndex { get; private set; }

        /// <summary>
        /// Пользовательский разделитель
        /// </summary>
        public string UserSeparator { get; private set; }

        /// <summary>
        /// Метод разделения стартового номера на части
        /// </summary>
        /// <returns>Массив, где 1й элемент - стартовый номер (подвергается +1), 2й элемент - индекс системы (не меняется), 3й - разделитель</returns>
        /// <exception cref="Exception">Ошибка, если номер не парситься</exception>
        private string[] SystemNumberSplit()
        {
            // Паттерн для поиска последней цифры
            string pattern = @"([\W_]+)(\d+)$";
            Match match = Regex.Match(SystemNumber, pattern);
            if (match.Success)
            {
                // Получаем символ, который отделяет части
                string separator = match.Groups[1].Value;
                // Получаем последнюю цифру
                string lastDigit = match.Groups[2].Value;
                // Все остальное до последней цифры
                string beforeLastDigit = SystemNumber.Substring(0, match.Index);

                return new string[] { lastDigit, beforeLastDigit, separator };
            }

            throw new Exception("Непонятный формат ввода стартового номера. Он должен всегда в конце содержать цифру, чтобы её можно было увеличивать на +1");
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
