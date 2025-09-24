using System;

namespace KPLN_Library_OpenDocHandler.Core
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
