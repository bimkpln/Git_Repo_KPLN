using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace KPLN_Clashes_Ribbon.Forms
{
    public class TextInputDialog
    {
        private readonly TextInput _form;
        
        public TextInputDialog(Window parent, string header)
        {
            Value = null;
            _form = new TextInput(parent, header);
        }
        public static string Value { get; set; }
        
        public void ShowDialog()
        {
            _form.ShowDialog();
        }
        
        public bool IsConfirmed()
        {

            if (Value == null) { return false; }
            return true;
        }
        
        public string GetLastPickedValue()
        {
            return Value;
        }
    }
}
