using Autodesk.Revit.DB;
using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
{
    public class OZKDuctAccessoryEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly Guid _prefParamGuid = new Guid("3a473dae-91b8-46e7-9db6-8f3371f4a5d4");
        private readonly Guid _sufParamGuid = new Guid("7cdd0d4b-7252-45a8-bc26-d0e63d871fa9");
        private string _currentTitle;
        private string _currentExample;

        public OZKDuctAccessoryEntity(FamilySymbol famSymbol, FamilyInstance[] famInstances)
        {
            CurrentFamilySymbol = famSymbol;
            CurrentFamilyInstances = famInstances;
            PreffixParam = famSymbol.get_Parameter(_prefParamGuid);
            SuffixParam = famSymbol.get_Parameter(_sufParamGuid);

            CurrentTitle = $"{famSymbol.FamilyName}: {famSymbol.Name} ({CurrentFamilyInstances.Length} шт.)";

            string sizeExmpl = "100x100";
            IEnumerator conEnum = CurrentFamilyInstances.FirstOrDefault().MEPModel.ConnectorManager.Connectors.GetEnumerator();
            while (conEnum.MoveNext() 
                && conEnum.Current is Connector connector
                && connector.Shape != ConnectorProfileType.Invalid)
            {
                if (connector.Shape == ConnectorProfileType.Round)
                {
                    sizeExmpl = "Ø100";
                    break;
                }
            }
            CurrentExample = $"{PreffixParam.AsString()}{sizeExmpl}{SuffixParam.AsString()}";

        }

        #region Данные для анализа
        // Ссылка на текущий тип семейства
        public FamilySymbol CurrentFamilySymbol { get; private set; }

        // Коллекция всех FamilyInstance текущего FamilySymbol
        public FamilyInstance[] CurrentFamilyInstances { get; private set; }

        // Параметр приставки в марке - "ААВ_Часть марки_До размера"
        public Parameter PreffixParam { get; private set; }

        // Парамтер окончания в марке - "ААВ_Часть марки_После размера"
        public Parameter SuffixParam { get; private set; }
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
