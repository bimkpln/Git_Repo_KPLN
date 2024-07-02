﻿using Autodesk.Revit.DB;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_Publication.Common
{
    public class ListBoxElement : INotifyPropertyChanged
    {
        public View View { get; set; }
        private System.Windows.Visibility? _visibility;
        public System.Windows.Visibility? Visibility
        {
            get
            {
                if (_visibility == null)
                {
                    return System.Windows.Visibility.Visible;
                }
                return _visibility;
            }
            set
            {
                _visibility = value;
                NotifyPropertyChanged();
            }
        }
        private bool _isSelected = false;
        public bool IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                if (value == _isSelected)
                {
                    return;
                }
                _isSelected = value;
                NotifyPropertyChanged();
            }
        }
        private bool _isChecked = false;
        public bool IsChecked
        {
            get
            {
                return _isChecked;
            }
            set
            {
                if (value == _isChecked)
                {
                    return;
                }
                _isChecked = value;
                NotifyPropertyChanged();
            }
        }
        public SolidColorBrush Fill { get; private set; }
        public string Name { get; private set; }
        public ListBoxElement(View view, bool isChecked)
        {
            //Print("ListBoxElement", KPLN_Loader.Preferences.MessageType.System_OK);
            View = view;
            IsChecked = isChecked;
            if (view.GetType() == typeof(ViewSheet))
            {
                Name = string.Format("{0}: {1}", (view as ViewSheet).SheetNumber, (view as ViewSheet).Name);
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 225, 210, 30));
            }
            else
            {
                Name = view.Name;
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 70, 160, 225));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            //Print(string.Format("NotifyPropertyChanged (" + propertyName + ")"), KPLN_Loader.Preferences.MessageType.System_Regular);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
