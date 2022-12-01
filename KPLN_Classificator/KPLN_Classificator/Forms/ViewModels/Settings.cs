using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.Forms.ViewModels
{
    public class Settings
    {
        public ObservableCollection<ParamNameItem> paramNameItems { get; set; }
        public int instanceOrType { get; set; }

        public Settings()
        {
            this.paramNameItems = new ObservableCollection<ParamNameItem>();
        }

        public void addParamName(string paramName)
        {
            paramNameItems.Add(new ParamNameItem(paramName));
            refreshNumbersOfParams();
        }

        public void addParamNameByIndex(string paramName, int index)
        {
            paramNameItems.Insert(index, new ParamNameItem(paramName));
            refreshNumbersOfParams();
        }

        public bool removeParamName(ParamNameItem paramNameItem)
        {
            if (paramNameItems.Remove(paramNameItem))
            {
                refreshNumbersOfParams();
                return true;
            }
            return false;
        }
        private void refreshNumbersOfParams()
        {
            for (int i = 0; i < paramNameItems.Count; i++)
            {
                paramNameItems[i].paramNumber = i + 1;
            }
        }
    }
}
