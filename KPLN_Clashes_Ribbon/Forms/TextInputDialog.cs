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
        public static string Value { get; set; }
        private TextInput form { get; set; }
        public TextInputDialog(Window parent, string header)
        {
            Value = null;
            form = new TextInput(parent, header);
        }
        public void ShowDialog()
        {
            form.ShowDialog();
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
