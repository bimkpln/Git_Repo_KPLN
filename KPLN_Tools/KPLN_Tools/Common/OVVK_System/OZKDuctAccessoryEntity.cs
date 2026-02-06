using Autodesk.Revit.DB;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
{
    public class OZKDuctAccessoryEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _currentTitle;
        private string _currentExample;

        public OZKDuctAccessoryEntity(FamilySymbol famSymbol, IEnumerable<FamilyInstance> famInstances, string preffixParamData, string suffixParamData)
        {
            CurrentFamilySymbol = famSymbol;
            CurrentFamilyInstances = famInstances.ToArray();
            PreffixParamData = preffixParamData;
            SuffixParamData = suffixParamData;


            CurrentTitle = $"{famSymbol.FamilyName}: {famSymbol.Name} ({CurrentFamilyInstances.Length} шт.)";


            string sizeExmpl = "100x100";
            IEnumerator conEnum = CurrentFamilyInstances.FirstOrDefault().MEPModel.ConnectorManager.Connectors.GetEnumerator();
            while (conEnum.MoveNext() 
                && conEnum.Current is Connector connector)
            {
                if (connector.Shape == ConnectorProfileType.Round)
                {
                    sizeExmpl = "Ø100";
                    break;
                }
            }


            CurrentExample = $"{PreffixParamData}{sizeExmpl}{SuffixParamData}";
        }

        #region Данные для анализа
        /// <summary>
        /// Ссылка на текущий тип семейства
        /// </summary>
        public FamilySymbol CurrentFamilySymbol { get; private set; }

        /// <summary>
        /// Коллекция всех FamilyInstance текущего FamilySymbol
        /// </summary>
        public FamilyInstance[] CurrentFamilyInstances { get; private set; }

        /// <summary>
        /// Данные приставки в марке - "ААВ_Часть марки_До размера"
        /// </summary>
        public string PreffixParamData { get; private set; }

        /// <summary>
        /// Данные окончания в марке - "ААВ_Часть марки_После размера"
        /// </summary>
        public string SuffixParamData { get; private set; }

        #endregion

        #region Данные для формы
        public string CurrentTitle
        {
            get => _currentTitle;
            private set
            {
                if (_currentTitle != value)
                {
                    _currentTitle = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string CurrentExample
        {
            get => _currentExample;
            private set
            {
                if (_currentExample != value)
                {
                    _currentExample = value;
                    NotifyPropertyChanged();
                }
            }
        }
        #endregion

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
