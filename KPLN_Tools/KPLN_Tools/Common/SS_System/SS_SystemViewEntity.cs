using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.SS_System
{
    public class SS_SystemViewEntity : INotifyPropertyChanged
    {
        private static GraphicsStyleWrapper _selectedLineStyle;
        private static ElectricalSystemType _selectedSystemType;
        private static bool _isLineDraw = false;
        private static string _systemNumber;
        private static string _userSystemIndex;
        private static string _userSeparator;
        private static int _startNumber;

        public event PropertyChangedEventHandler PropertyChanged;

        public SS_SystemViewEntity(Dictionary<string, GraphicsStyleWrapper> lineStyles, Dictionary<string, ElectricalSystemType> electricalSystemTypes)
        {
            LineStyles = lineStyles;
            ElectricalSystemTypes = electricalSystemTypes;

            if(SelectedSystemType == ElectricalSystemType.UndefinedSystemType)
                SelectedSystemType = ElectricalSystemTypes.FirstOrDefault().Value;

            if (string.IsNullOrEmpty(_systemNumber))
                SystemNumber = "1.1";
        }

        /// <summary>
        /// Коллекция типов линий в проекте
        /// </summary>
        public Dictionary<string, GraphicsStyleWrapper> LineStyles { get; set; }

        /// <summary>
        /// Коллекция типов систем СС
        /// </summary>
        public Dictionary<string, ElectricalSystemType> ElectricalSystemTypes { get; set; }

        /// <summary>
        /// Выбранный тип линии для построения в проекте
        /// </summary>
        public GraphicsStyleWrapper SelectedLineStyle
        {
            get => _selectedLineStyle;
            set
            {
                if (_selectedLineStyle != value)
                {
                    _selectedLineStyle = value;
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

                    if (_isLineDraw)
                        SelectedLineStyle = LineStyles.FirstOrDefault().Value;
                    else
                        SelectedLineStyle = null;
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

                    string[] splitedNumber = NumberService.SystemNumberSplit(_systemNumber);
                    UserSystemIndex = splitedNumber[1];
                    UserSeparator = splitedNumber[2];
                    
                    StartNumber = int.Parse(splitedNumber[0]);
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

                    SystemNumber = string.Format("{0}{1}{2}", UserSeparator, UserSystemIndex, StartNumber);
                }
            }
        }

        /// <summary>
        /// Пользовательский индекс системы
        /// </summary>
        public string UserSystemIndex 
        {
            get => _userSystemIndex;
            private set
            {
                _userSeparator = value;
            }
        }

        /// <summary>
        /// Пользовательский разделитель
        /// </summary>
        public string UserSeparator 
        {
            get => _userSeparator;
            private set
            {
                _userSystemIndex = value;
            }
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
