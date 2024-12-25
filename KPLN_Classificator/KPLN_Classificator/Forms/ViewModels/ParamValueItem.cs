using KPLN_Classificator.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.Forms.ViewModels
{
    public class ParamValueItem : BaseViewModel
    {
        private string _colourOfValueBack { get; set; }
        public string colourOfValueBack
        {
            get
            {
                return _colourOfValueBack;
            }
            set
            {
                _colourOfValueBack = value;
                NotifyPropertyChanged();
            }
        }
        private string _paramValue { get; set; }
        public string paramValue
        {
            get
            {
                return _paramValue;
            }
            set
            {
                _paramValue = value;
                NotifyPropertyChanged();
                parent?.checkChanges();
                
                if (_paramValue != null && !isParamValueCorrect(_paramValue)) colourOfValueBack = ColourUtils.WRONG_VALUE;
                else colourOfValueBack = ColourUtils.CORRECT_VALUE;
            }
        }
        private string _visibleOfFindButton { get; set; } = "Hidden";
        public RuleItem parent { get; set; }
        private int _valueNumber { get; set; }
        public int valueNumber
        {
            get
            {
                return _valueNumber;
            }
            set
            {
                _valueNumber = value;
                NotifyPropertyChanged();
            }
        }

        public ParamValueItem(string paramValue, RuleItem parent)
        {
            this.paramValue = paramValue;
            this.parent = parent;
            this.colourOfValueBack = ColourUtils.CORRECT_VALUE;
        }

        public override bool Equals(object obj)
        {
            return obj is ParamValueItem item &&
                   paramValue == item.paramValue;
        }

        private bool isParamValueCorrect(string paramValue)
        {
            Stack<char> stack = new Stack<char>();
            foreach (var item in paramValue.ToCharArray())
            {
                if (item == '[')
                {
                    stack.Push(item);
                }
                else if (item == ']' && stack.Count != 0 && stack.Peek() == '[')
                {
                    stack.Pop();
                }
                else if (item == ']')
                {
                    stack.Push(item);
                }
            }
            return stack.Count == 0;
        }

        public override int GetHashCode()
        {
            int hashCode = 663154009;
            hashCode = hashCode * 31 + EqualityComparer<string>.Default.GetHashCode(paramValue);
            return hashCode;
        }

        public override string ToString()
        {
            return this.GetType().ToString() + " " + paramValue;
        }
    }
}
