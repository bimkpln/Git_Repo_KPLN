﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Parameters_Ribbon.Common.CopyElemParamData
{
    public class ListBoxElement
    {
        public ObservableCollection<ListBoxElement> SubElements = new ObservableCollection<ListBoxElement>();
        
        public object Data { get; set; }
        
        public string Name { get; private set; }
        
        public string ToolTip { get; private set; }
        
        public ListBoxElement(object data, string name)
        {
            Data = data;
            Name = name;
        }
        
        public ListBoxElement(object data, string name, string toolTip) : this(data, name)
        {
            ToolTip = toolTip;
        }
        
    }
}
