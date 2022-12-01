using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.Forms.ViewModels
{
    public class ParamNameItem : BaseViewModel
    {
        private string _paramName { get; set; }
        public string paramName
        {
            get
            {
                return _paramName;
            }
            set
            {
                _paramName = value;
                NotifyPropertyChanged();
            }
        }
        private int _paramNumber { get; set; }
        public int paramNumber
        {
            get
            {
                return _paramNumber;
            }
            set
            {
                _paramNumber = value;
                NotifyPropertyChanged();
            }
        }
        public ParamNameItem(string paramName)
        {
            this.paramName = paramName;
        }

        public override bool Equals(object obj)
        {
            return obj is ParamNameItem item &&
                   paramName == item.paramName &&
                   paramNumber == item.paramNumber;
        }

        public override int GetHashCode()
        {
            int hashCode = -946399840;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(paramName);
            hashCode = hashCode * -1521134295 + paramNumber.GetHashCode();
            return hashCode;
        }
    }
}
