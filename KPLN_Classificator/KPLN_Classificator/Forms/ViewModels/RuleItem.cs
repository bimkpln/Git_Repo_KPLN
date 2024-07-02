using KPLN_Classificator.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BuiltInCategory = Autodesk.Revit.DB.BuiltInCategory;

namespace KPLN_Classificator.Forms.ViewModels
{
    public class RuleItem : BaseViewModel
    {
        private long hashSumOfRuleItem;
        private string _colourOfRule { get; set; }
        public string colourOfRule
        {
            get
            {
                return _colourOfRule;
            }
            set
            {
                _colourOfRule = value;
                NotifyPropertyChanged();
            }
        }
        private string _familyNameValue { get; set; }
        public string familyNameValue
        {
            get
            {
                return _familyNameValue;
            }
            set
            {
                _familyNameValue = value;
                NotifyPropertyChanged();
                checkChanges();
            }
        }
        private string _familyNameTextBox { get; set; }
        public string familyNameTextBox
        {
            get
            {
                return _familyNameTextBox;
            }
            set
            {
                _familyNameTextBox = value;
                NotifyPropertyChanged();
            }
        }
        private string _typeNameValue { get; set; }
        public string typeNameValue
        {
            get
            {
                return _typeNameValue;
            }
            set
            {
                _typeNameValue = value;
                NotifyPropertyChanged();
                checkChanges();
            }
        }
        private string _typeNameTextBox { get; set; }

        public string typeNameTextBox
        {
            get
            {
                return _typeNameTextBox;
            }
            set
            {
                _typeNameTextBox = value;
                NotifyPropertyChanged();
            }
        }

        private string _parameterName { get; set; }

        public string parameterName
        {
            get
            {
                return _parameterName;
            }
            set
            {
                _parameterName = value;
                NotifyPropertyChanged();
                checkChanges();
            }
        }

        private string _parameterValue { get; set; }

        public string parameterValue
        {
            get
            {
                return _parameterValue;
            }
            set
            {
                _parameterValue = value;
                NotifyPropertyChanged();
                checkChanges();
            }
        }


        private BuiltInCategory _builtInCategoryName { get; set; }

        public BuiltInCategory builtInCategoryName
        {
            get
            {
                return _builtInCategoryName;
            }
            set
            {
                _builtInCategoryName = value;
                familyNameTextBox = builtInCategoryName.Equals(BuiltInCategory.OST_Rooms) ? "Параметр \"Назначение\" помещения содержит:" : "Имя семейства содержит:";
                typeNameTextBox = builtInCategoryName.Equals(BuiltInCategory.OST_Rooms) ? "Параметр \"Имя\" помещения содержит:" : "Имя типа содержит:";
                NotifyPropertyChanged();
                checkChanges();
            }
        }

        private int _ruleNumber { get; set; }
        public int ruleNumber
        {
            get
            {
                return _ruleNumber;
            }
            set
            {
                _ruleNumber = value;
                NotifyPropertyChanged();
            }
        }
        public ObservableCollection<ParamValueItem> valuesOfParams { get; set; }
        public static List<BuiltInCategory> builtInCats { get; set; }
        private int _elemsCountInModel { get; set; }
        public int elemsCountInModel
        {
            get
            {
                return _elemsCountInModel;
            }
            set
            {
                _elemsCountInModel = value;
                NotifyPropertyChanged();
            }
        }

        public RuleItem(string familyNameValue, string typeNameValue, string parameterName, string parameterValue, BuiltInCategory builtInCategoryName)
        {
            this.familyNameValue = familyNameValue;
            this.typeNameValue = typeNameValue;
            this.parameterName = parameterName;
            this.parameterValue = parameterValue;
            this.valuesOfParams = new ObservableCollection<ParamValueItem>();
            this.builtInCategoryName = builtInCategoryName;
            this.elemsCountInModel = -1;
            hashSumOfRuleItem = getHashSumOfRuleItem();
        }

        public static void refreshNumbersOfValues(ObservableCollection<ParamValueItem> valuesOfParams)
        {
            for (int i = 0; i < valuesOfParams.Count; i++)
            {
                valuesOfParams[i].valueNumber = i + 1;
            }
        }

        public void addValueOfParam(string paramValue)
        {
            valuesOfParams.Add(new ParamValueItem(paramValue, this));
            checkChanges();
            refreshNumbersOfValues(valuesOfParams);
        }

        public void addValueOfParamByIndex(string paramValue, int index)
        {
            valuesOfParams.Insert(index, new ParamValueItem(paramValue, this));
            checkChanges();
            refreshNumbersOfValues(valuesOfParams);
        }

        public bool removeValueOfParam(ParamValueItem valueItem)
        {
            if (valuesOfParams.Remove(valueItem))
            {
                checkChanges();
                refreshNumbersOfValues(valuesOfParams);
                return true;
            }
            return false;
        }

        public override bool Equals(object obj)
        {
            return obj is RuleItem item &&
                   familyNameValue == item.familyNameValue &&
                   typeNameValue == item.typeNameValue &&
                   builtInCategoryName == item.builtInCategoryName;
        }

        public override int GetHashCode()
        {
            int hashCode = -25894503;
            hashCode = hashCode * 31 + EqualityComparer<string>.Default.GetHashCode(familyNameValue);
            hashCode = hashCode * 31 + EqualityComparer<string>.Default.GetHashCode(typeNameValue);
            hashCode = hashCode * 31 + EqualityComparer<string>.Default.GetHashCode(parameterName);
            hashCode = hashCode * 31 + EqualityComparer<string>.Default.GetHashCode(parameterValue);
            hashCode = hashCode * 31 + builtInCategoryName.GetHashCode();
            return hashCode;
        }

        public override string ToString()
        {
            return this.GetType().ToString() + " " + familyNameValue + " " + typeNameValue;
        }

        private long getHashSumOfRuleItem()
        {
            if (valuesOfParams == null)
                return hashSumOfRuleItem;
            else
            {
                long sum = 0L;
                sum += GetHashCode();
                foreach (ParamValueItem valueItem in valuesOfParams)
                {
                    sum += valueItem.GetHashCode();
                }
                return sum;
            }
        }

        public void setHashSum()
        {
            hashSumOfRuleItem = getHashSumOfRuleItem();
        }

        public void checkChanges()
        {
            if (hashSumOfRuleItem != getHashSumOfRuleItem())
            {
                colourOfRule = ColourUtils.NEW_RULE;
            }
            else
            {
                colourOfRule = ColourUtils.OLD_RULE;
            }
        }
    }
}