﻿using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Loader.Forms.Common
{
    /// <summary>
    /// Описание события для окна LoaderStatusForm
    /// </summary>
    internal sealed class LoaderEvantEntity : INotifyPropertyChanged
    {
        private System.Windows.Media.Brush _loadColor;
        private string _loadDescription;

        public LoaderEvantEntity(string descr)
        {
            _loadDescription = descr;
        }

        public string LoadDescription
        {
            get => _loadDescription;
            set
            {
                _loadDescription = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Media.Brush LoadColor
        {
            get => _loadColor;
            set
            {
                _loadColor = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
