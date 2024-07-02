using System;

namespace KPLN_BIMTools_Ribbon.Common
{
    public class FieldChangedEventArgs : EventArgs
    {
        public string NewValue { get; }

        public FieldChangedEventArgs(string newValue)
        {
            NewValue = newValue;
        }
    }
}
