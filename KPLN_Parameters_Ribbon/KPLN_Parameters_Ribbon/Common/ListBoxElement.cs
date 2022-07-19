using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Parameters_Ribbon.Common
{
    public class ListBoxElement : INotifyPropertyChanged
    {
        public ObservableCollection<ListBoxElement> SubElements = new ObservableCollection<ListBoxElement>();
        public object Data { get; set; }
        public string Name { get; private set; }
        public string ToolTip { get; private set; }
        public ListBoxElement(object data, string name, string toolTip)
        {
            Data = data;
            Name = name;
            ToolTip = toolTip;
        }
        public ListBoxElement(object data, string name)
        {
            Data = data;
            Name = name;
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
